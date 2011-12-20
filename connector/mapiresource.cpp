/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011 Shaheed Haque <srhaque@theiet.org>.
 *
 * Akonadi Exchange Resource is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * Akonadi Exchange Resource is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with Akonadi Exchange Resource.
 * If not, see <http://www.gnu.org/licenses/>.
 */

#include "mapiresource.h"

#include <QtDBus/QDBusConnection>

#include <KConfigGroup>
#include <KLocalizedString>
#include <KWindowSystem>
#include <KStandardDirs>

#include <Akonadi/AgentManager>
#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>
#include <akonadi/kmime/messageparts.h>
#include <kmime/kmime_message.h>

#include "mapiconnector2.h"

/**
 * Display ids in the same format we use when stored in Akonadi.
 */
#define ID_BASE 36

using namespace Akonadi;

/**
 * We store all objects in Akonadi using the densest string representation to hand.
 */
static QString toStringId(qulonglong id)
{
	return QString::number(id, ID_BASE);
}

static qulonglong fromStringId(const QString &id)
{
	return id.toULongLong(0, ID_BASE);
}

/**
 * We store all objects in Akonadi using the densest string representation to hand.
 */
const QChar FullId::fidIdSeparator = QChar::fromAscii('/');

FullId::FullId(qulonglong f, qulonglong s)
{
	first = f;
	second = s;
}

FullId::FullId(const QString &id)
{
	int separator = id.indexOf(fidIdSeparator);
	first = fromStringId(id.left(separator));
	second = fromStringId(id.mid(separator + 1));
}

QString FullId::toString() const
{
	return toStringId(first).append(fidIdSeparator).append(toStringId(second));
}

MapiResource::MapiResource(const QString &id, const QString &desktopName, const char *mapiFolderFilter, const char *mapiMessageType, const QString &itemMimeType) :
	ResourceBase(id),
	m_mapiFolderFilter(QString::fromAscii(mapiFolderFilter)),
	m_mapiMessageType(QString::fromAscii(mapiMessageType)),
	m_itemMimeType(itemMimeType),
	m_connection(new MapiConnector2()),
	m_connected(false)
{
	if (name() == identifier()) {
		setName(desktopName);
	}

	setHierarchicalRemoteIdentifiersEnabled(true);
	//setCollectionStreamingEnabled(true);
	//setItemStreamingEnabled(true);
}

MapiResource::~MapiResource()
{
	logoff();
	delete m_connection;
}

void MapiResource::error(const QString &message)
{
	kError() << message;
	emit status(Broken, message);
	cancelTask(message);
}

void MapiResource::error(const MapiFolder &folder, const QString &body)
{
	static QString prefix = QString::fromAscii("Error %1: %2");
	QString message = prefix.arg(toStringId(folder.id())).arg(body);

	error(message);
}

void MapiResource::error(const Akonadi::Collection &collection, const QString &body)
{
	static QString prefix = QString::fromAscii("Error %1(%2): %3");
	QString message = prefix.arg(collection.remoteId()).arg(collection.name()).arg(body);

	error(message);
}

void MapiResource::error(const MapiMessage &msg, const QString &body)
{
	static QString prefix = QString::fromAscii("Error %1/%2: %3");
	QString message = prefix.arg(toStringId(msg.folderId())).arg(toStringId(msg.id())).arg(body);

	error(message);
}

void MapiResource::fetchCollections(MapiDefaultFolder rootFolder, Akonadi::Collection::List &collections)
{
	kDebug() << "fetch all collections";

	if (!logon()) {
		// Come back later.
		deferTask();
		return;
	}

	// create the new collection
	mapi_id_t rootId;
	if (!m_connection->defaultFolder(rootFolder, &rootId))
	{
		error(i18n("cannot find Exchange folder root"));
		return;
	}
	Collection root;
	FullId remoteId(0, rootId);
	QStringList contentTypes;

	contentTypes << m_itemMimeType << Akonadi::Collection::mimeType();
	root.setName(name());
	root.setRemoteId(remoteId.toString());
	root.setParentCollection(Collection::root());
	root.setContentMimeTypes(contentTypes);
	root.setRights(Akonadi::Collection::ReadOnly);
	collections.append(root);
	fetchCollections(root.name(), root, collections);
	emit status(Running, i18n("fetched collections: %1").arg(collections.size()));
}

void MapiResource::fetchCollections(const QString &path, const Collection &parent, Akonadi::Collection::List &collections)
{
	kDebug() << "fetch collections in:" << path;

	FullId parentRemoteId(parent.remoteId());
	MapiFolder parentFolder(m_connection, "MapiResource::retrieveCollection", parentRemoteId.second);
	if (!parentFolder.open()) {
		error(parentFolder, i18n("cannot open Exchange folder list"));
		return;
	}

	QList<MapiFolder *> list;
	emit status(Running, i18n("fetching folder list from Exchange: %1").arg(path));
	if (!parentFolder.childrenPull(list, m_mapiFolderFilter)) {
		error(parentFolder, i18n("cannot fetch folder list from Exchange"));
		return;
	}

	QChar separator = QChar::fromAscii('/');
	QStringList contentTypes;

	contentTypes << m_itemMimeType << Akonadi::Collection::mimeType();
	foreach (MapiFolder *data, list) {
		Collection child;
		FullId remoteId(parentRemoteId.second, data->id());

		child.setName(data->name);
		child.setRemoteId(remoteId.toString());
		child.setParentCollection(parent);
		child.setContentMimeTypes(contentTypes);
		collections.append(child);

		// Recurse down...
		QString currentPath = path + separator + child.name();
		fetchCollections(currentPath, child, collections);
		delete data;
	}
}

void MapiResource::fetchItems(const Akonadi::Collection &collection, Item::List &items, Item::List &deletedItems)
{
	kDebug() << "fetch items from collection:" << collection.name();

	if (!logon()) {
		// Come back later.
		deferTask();
		return;
	}

	// Find all item that are already in this collection in Akonadi.
	QSet<FullId> knownRemoteIds;
	QMap<FullId, Item> knownItems;
	{
		emit status(Running, i18n("Fetching items from Akonadi cache"));
		ItemFetchJob *fetch = new ItemFetchJob( collection );

		Akonadi::ItemFetchScope scope;
		// we are only interested in the items from the cache
		scope.setCacheOnly(true);
		// we don't need the payload (we are mainly interested in the remoteID and the modification time)
		scope.fetchFullPayload(false);
		fetch->setFetchScope(scope);
		if ( !fetch->exec() ) {
			error(collection, i18n("unable to fetch listing of collection: %1", fetch->errorString()));
			return;
		}
		Item::List existingItems = fetch->items();
		foreach (Item item, existingItems) {
			// store all the items that we already know
			FullId remoteId(item.remoteId());
			knownRemoteIds.insert(remoteId);
			knownItems.insert(remoteId, item);
		}
	}
	kError() << "knownRemoteIds:" << knownRemoteIds.size();

	FullId parentRemoteId(collection.remoteId());
	MapiFolder parentFolder(m_connection, "MapiResource::retrieveItems", parentRemoteId.second);
	if (!parentFolder.open()) {
		error(collection, i18n("unable to open collection"));
		return;
	}

	// Get the folder content for the collection.
	QList<MapiItem *> list;
	emit status(Running, i18n("Fetching collection: %1", collection.name()));
	if (!parentFolder.childrenPull(list)) {
		error(collection, i18n("unable to fetch collection"));
		return;
	}
	kError() << "fetched:" << list.size() << "items from collection:" << collection.name();

	QSet<FullId> checkedRemoteIds;
	// run though all the found data...
	foreach (MapiItem *data, list) {
		FullId remoteId(parentRemoteId.second, data->id());
		checkedRemoteIds << remoteId; // store for later use

		if (!knownRemoteIds.contains(remoteId)) {
			// we do not know this remoteID -> create a new empty item for it
			Item item(m_itemMimeType);
			item.setParentCollection(collection);
			item.setRemoteId(remoteId.toString());
			item.setRemoteRevision(QString::number(1));
			//item.setModificationTime(data->modified());
			items << item;
		} else {
			// this item is already known, check if it was update in the meanwhile
			Item& existingItem = knownItems.find(remoteId).value();
// 				kDebug() << "Item("<<existingItem.id()<<":"<<data.id<<":"<<existingItem.revision()<<") is already known [Cache-ModTime:"<<existingItem.modificationTime()
// 						<<" Server-ModTime:"<<data.modified<<"] Flags:"<<existingItem.flags()<<"Attrib:"<<existingItem.attributes();
			if (existingItem.modificationTime() < data->modified()) {
				kDebug() << existingItem.id()<<"=> this item has changed";

				// force akonadi to call retrieveItem() for this item in order to get updated data
				int revision = existingItem.remoteRevision().toInt();
				existingItem.clearPayload();
				existingItem.setRemoteRevision( QString::number(++revision) );
				items << existingItem;
			}
		}
		delete data;
		// TODO just for debugging...
		if (items.size() > 3) {
			//break;
		}
	}

	// now check if some of the items need to be removed
	knownRemoteIds.subtract(checkedRemoteIds);

	foreach (const FullId &remoteId, knownRemoteIds) {
		deletedItems << knownItems.value(remoteId);
	}

	foreach(Item item, items) {
		kDebug() << "[Item-Dump] ID:"<<item.id()<<"RemoteId:"<<item.remoteId()<<"Revision:"<<item.revision()<<"ModTime:"<<item.modificationTime();
	}

	// We fetched a load of stuff. This seems like a good place to force 
	// any subsequent activity to re-attempt the login.
	logoff();
}

bool MapiResource::logon(void)
{
	if (!m_connected) {
		// logon to exchange (if needed)
		emit status(Running, i18n("Logging in to Exchange as %1").arg(m_profileName));
		m_connected = m_connection->login(m_profileName);
	}
	if (!m_connected) {
		emit status(Broken, i18n("Unable to login as %1").arg(m_profileName));
	}
	return m_connected;
}

void MapiResource::logoff(void)
{
	// There is no logoff operation. We just want to make sure we retry the
	// logon next time.
	m_connected = false;
}

void MapiResource::profileSet(const QString &profileName)
{
	m_profileName = profileName;
}

#include "mapiresource.moc"
