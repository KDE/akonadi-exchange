/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011 Robert Gruber <rgruber@users.sourceforge.net>, Shaheed Haque
 * <srhaque@theiet.org>.
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

#include "settings.h"
#include "settingsadaptor.h"

#include "exgalresource.h"

#include <QtDBus/QDBusConnection>

#include <KLocalizedString>
#include <KABC/Addressee>
#include <KWindowSystem>

#include "mapiconnector2.h"
#include "profiledialog.h"

using namespace Akonadi;

ExGalResource::ExGalResource(const QString &id) : 
	MapiResource(id, i18n("Exchange Contacts"), IPF_CONTACT, "IPM.Contact", QString::fromAscii("text/directory")),
	m_galId(0, 0),
	m_begin(true)
{
	new SettingsAdaptor(Settings::self());
	QDBusConnection::sessionBus().registerObject(QLatin1String("/Settings"),
						     Settings::self(), 
						     QDBusConnection::ExportAdaptors);
	//setItemStreamingEnabled(true);
}

ExGalResource::~ExGalResource()
{
}

void ExGalResource::retrieveCollections()
{
	// First select who to log in as.
	profileSet(Settings::self()->profileName());

	// We are going to return both the user's contacts as well as the GAL.
	// First, the GAL, then the user's contacts...
	Collection::List collections;
	Collection gal;

	gal.setName(i18n("Exchange Global Address List"));
	gal.setRemoteId(m_galId.toString());
	gal.setParentCollection(Collection::root());
	gal.setContentMimeTypes(QStringList(m_itemMimeType));
	gal.setRights(Akonadi::Collection::ReadOnly);
	collections.append(gal);
	fetchCollections(Contacts, collections);

	// Notify Akonadi about the new collections.
	collectionsRetrieved(collections);
}

void ExGalResource::retrieveItems(const Akonadi::Collection &collection)
{
	Item::List items;
	Item::List deletedItems;

	if (collection.remoteId() != m_galId.toString()) {
		// This request is NOT for the GAL.
		fetchItems(collection, items, deletedItems);
		itemsRetrievedIncremental(items, deletedItems);
		itemsRetrievalDone();
	} else {
		unsigned requestedCount = 100;

		if (!m_begin) {
			kError() << "TEMP DEBUG ONLY All done!!!!!!!!!!!!!!!!!!!!!!!";
			itemsRetrievalDone();
		}

		QList<GalMember> list;
		emit status(Running, i18n("Fetching GAL from Exchange"));
		if (!m_connection->fetchGAL(m_begin, requestedCount, list)) {
			// Error. Next time, start over.
			error(i18n("cannot fetch GAL from Exchange"));
			m_begin = true;
			return;
		}
		foreach (GalMember data, list) {
			Item item(m_itemMimeType);
			item.setParentCollection(collection);
			// find a better remote id (but entryid seams to be binary)
			//item.setRemoteId(data.id);
			item.setRemoteId(data.email);
			item.setRemoteRevision(QString::number(1));

			// prepare payload
			KABC::Addressee addressee;
			addressee.setName(data.name);
			addressee.setNickName(data.nick);
			addressee.setOrganization(data.organization);

			QStringList emails;
			emails << data.email;
			addressee.setEmails(emails);

			item.setPayload<KABC::Addressee>(addressee);
			items << item;
		}
		itemsRetrievedIncremental(items, deletedItems);
		if ((unsigned)(items.size() + deletedItems.size()) < requestedCount) {
			// All done. Next time, start over.
			itemsRetrievalDone();
			m_begin = true;
		}
	}
	kError() << "new/changed items:" << items.size() << "deleted items:" << deletedItems.size();
}

bool ExGalResource::retrieveItem( const Akonadi::Item &item, const QSet<QByteArray> &parts )
{
  Q_UNUSED( item );
  Q_UNUSED( parts );

  // TODO: this method is called when Akonadi wants more data for a given item.
  // You can only provide the parts that have been requested but you are allowed
  // to provide all in one go

  return true;
}

void ExGalResource::aboutToQuit()
{
  // TODO: any cleanup you need to do while there is still an active
  // event loop. The resource will terminate after this method returns
}

void ExGalResource::configure( WId windowId )
{
	ProfileDialog dlgConfig(Settings::self()->profileName());
  	if (windowId)
		KWindowSystem::setMainWindow(&dlgConfig, windowId);

	if (dlgConfig.exec() == KDialog::Accepted) {

		QString profile = dlgConfig.getProfileName();
		Settings::self()->setProfileName( profile );
		Settings::self()->writeConfig();

		synchronize();
		emit configurationDialogAccepted();
	} else {
		emit configurationDialogRejected();
	}
}

void ExGalResource::itemAdded( const Akonadi::Item &item, const Akonadi::Collection &collection )
{
  Q_UNUSED( item );
  Q_UNUSED( collection );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has created an item in a collection managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

void ExGalResource::itemChanged( const Akonadi::Item &item, const QSet<QByteArray> &parts )
{
  Q_UNUSED( item );
  Q_UNUSED( parts );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has changed an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

void ExGalResource::itemRemoved( const Akonadi::Item &item )
{
  Q_UNUSED( item );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has deleted an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

AKONADI_RESOURCE_MAIN(ExGalResource)

#include "exgalresource.moc"
