/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011 Robert Gruber <rgruber@users.sourceforge.net>
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

#include "excalresource.h"

#include "settings.h"
#include "settingsadaptor.h"

#include <QtDBus/QDBusConnection>

#include <KLocalizedString>
#include <KWindowSystem>

#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>

#include <KCal/Event>

#include "mapiconnector2.h"
#include "profiledialog.h"

using namespace Akonadi;

ExCalResource::ExCalResource( const QString &id )
  : ResourceBase( id ),
	m_connection(new MapiConnector2()),
	m_connected(false)
{
	new SettingsAdaptor( Settings::self() );
	QDBusConnection::sessionBus().registerObject( QLatin1String( "/Settings" ),
							Settings::self(), QDBusConnection::ExportAdaptors );
}

ExCalResource::~ExCalResource()
{
	logoff();
	delete m_connection;
}

void ExCalResource::retrieveCollections()
{
	kDebug() << "retrieveCollections() called";

	if (!logon()) {
		return;
	}

	QStringList folderMimeType;
	folderMimeType << QString::fromAscii("text/calendar");

	// create the new collection
	Collection root;
	root.setParentCollection(Collection::root());
	root.setContentMimeTypes(folderMimeType);

	MapiFolder rootFolder(m_connection, "ExCalResource::retrieveCollections", 0);
	if (!rootFolder.open()) {
		return;
	}

	QList<MapiFolder *> list;
	emit status(Running, i18n("Fetching folder list from Exchange"));
	if (!rootFolder.childrenPull(list, QString::fromAscii(IPF_APPOINTMENT))) {
		kError() << "cannot open root folder";
		return;
	}
	if (list.size() == 0) {
		kError() << "no calendar folders found";
		emit status(Broken, i18n("No calendar folders in Exchange"));
		return;
	}

	Collection::List collections;
	bool done = false;
	foreach (MapiFolder *data, list) {
		// TODO: take the first calender for now, but Exchange might have more calendar folders
		if (!done) {
			root.setRemoteId(data->id());
			root.setName(i18n("Exchange: ") + data->name);
			done = true;
		}
		kError() << "delete " << data;
		delete data;
	}
	collections.append(root);

	// notify akonadi about the new calendar collection
	collectionsRetrieved(collections);
}

void ExCalResource::retrieveItems( const Akonadi::Collection &collection )
{
	kDebug() << "retrieveItems() called for collection "<< collection.id();

	if (!logon()) {
		return;
	}

	// find all item that are already in this collection
	QSet<QString> knownRemoteIds;
	QMap<QString, Item> knownItems;
	{
		emit status(Running, i18n("Feching items from Akonadi cache"));
		ItemFetchJob *fetch = new ItemFetchJob( collection );

		Akonadi::ItemFetchScope scope;
		// we are only interested in the items from the cache
		scope.setCacheOnly(true);
		// we don't need the payload (we are mainly interested in the remoteID and the modification time)
		scope.fetchFullPayload(false);
		fetch->setFetchScope(scope);
		if ( !fetch->exec() ) {
			emit status(Broken, i18n("Unable to fetch listing of collection '%1': %2" , collection.name() ,fetch->errorString()) );
			return;
		}
		Item::List existingItems = fetch->items();
		foreach (Item item, existingItems) {
			// store all the items that we already know
			knownRemoteIds.insert( item.remoteId() );
			knownItems.insert(item.remoteId(), item);
		}
	}

	MapiFolder parentFolder(m_connection, "ExCalResource::retrieveItems", collection.remoteId().toULongLong());
	if (!parentFolder.open()) {
		kError() << "open failed!";
		emit status(Broken, i18n("Unable to open collection: %1", collection.name()));
		return;
	}

	// get the folder content for the collection
	Item::List items;
	QList<MapiItem *> list;
	emit status(Running, i18n("Fetching collection: %1", collection.name()));
	if (!parentFolder.childrenPull(list)) {
		kError() << "childrenPull failed!";
		emit status(Broken, i18n("Unable to fetch collection: %1", collection.name()));
		return;
	}
	kDebug() << "size of collection" << collection.id() << "is" << list.size();

	QSet<QString> checkedRemoteIds;
	// run though all the found data...
	foreach (const MapiItem *data, list) {

		checkedRemoteIds << data->id; // store for later use

		if (!knownRemoteIds.contains(data->id)) {
			// we do not know this remoteID -> create a new empty item for it
			Item item(QString::fromAscii("text/calendar"));
			item.setParentCollection(collection);
			item.setRemoteId(data->id);
			item.setRemoteRevision(QString::number(1));
			items << item;
		} else {
			// this item is already known, check if it was update in the meanwhile
			Item& existingItem = knownItems.find(data->id).value();
// 				kDebug() << "Item("<<existingItem.id()<<":"<<data.id<<":"<<existingItem.revision()<<") is already known [Cache-ModTime:"<<existingItem.modificationTime()
// 						<<" Server-ModTime:"<<data.modified<<"] Flags:"<<existingItem.flags()<<"Attrib:"<<existingItem.attributes();
			if (existingItem.modificationTime() < data->modified) {
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
// 			if (items.size() > 3) {
// 				break;
// 			}
	}

	// now check if some of the items need to be removed
	knownRemoteIds.subtract(checkedRemoteIds);

	Item::List deletedItems;
	foreach (const QString &remoteId, knownRemoteIds) {
		deletedItems << knownItems.value(remoteId);
	}

	kDebug() << "itemsRetrievedIncremental(): fetched"<<items.size()<<"new/changed items; delete"<<deletedItems.size()<<"items";
	itemsRetrievedIncremental(items, deletedItems);

	foreach(Item item, items) {
		kDebug() << "[Item-Dump] ID:"<<item.id()<<"RemoteId:"<<item.remoteId()<<"Revision:"<<item.revision()<<"ModTime:"<<item.modificationTime();
	}

	// We fetched a load of stuff. This seems like a good place to force 
	// any subsequent activity to re-attempt the login.
	logoff();
}

bool ExCalResource::retrieveItem( const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts )
{
	Q_UNUSED( parts );

	kDebug() << "retrieveItem() called for item "<< itemOrig.id() << "remoteId:" << itemOrig.remoteId();

	if (!logon()) {
		return false;
	}

	qulonglong messageId = itemOrig.remoteId().toULongLong();
	qulonglong folderId = currentCollection().remoteId().toULongLong();
	MapiAppointment message(m_connection, "ExCalResource::retrieveItem", messageId);
	if (!message.open(folderId)) {
		kError() << "open failed!";
		emit status(Broken, i18n("Unable to open item: { %1, %2 }", currentCollection().name(), messageId));
		return false;
	}

	// find the remoteId of the item and the collection and try to fetch the needed data from the server
	kWarning() << "fetching item: {" <<
		currentCollection().name() << "," << itemOrig.id() << "} = {" <<
		folderId << "," << messageId << "}";
	emit status(Running, i18n("Fetching item: { %1, %2 }", currentCollection().name(), messageId));
	if (!message.propertiesPull()) {
		kError() << "propertiesPull failed!";
		emit status(Broken, i18n("Unable to fetch item: { %1, %2 }", currentCollection().name(), messageId));
		return false;
	}
	kDebug() << "got message; item:"<<message.id<<":"<<message.title;

	// Create a clone of the passed in Item and fill it with the payload
	Akonadi::Item item(itemOrig);

	KCal::Event* event = new KCal::Event;
	event->setUid(item.remoteId());
	event->setSummary( message.title );
	event->setDtStart( KDateTime(message.begin) );
	event->setDtEnd( KDateTime(message.end) );
	event->setCreated( KDateTime(message.created) );
	event->setLastModified( KDateTime(message.modified) );
	event->setDescription( message.text );
	event->setOrganizer( message.sender );
	event->setLocation( message.location );
	if (message.reminderActive) {
		KCal::Alarm* alarm = new KCal::Alarm( dynamic_cast<KCal::Incidence*>( event ) );
		// TODO Maybe we should check which one is set and then use either the time or the delte
		// KDateTime reminder(message.reminderTime);
		// reminder.setTimeSpec( KDateTime::Spec(KDateTime::UTC) );
		// alarm->setTime( reminder );
		alarm->setStartOffset(KCal::Duration(message.reminderMinutes * -60));
		alarm->setEnabled( true );
		event->addAlarm( alarm );
	}

	foreach (Attendee att, message.attendees) {
		if (att.isOrganizer()) {
			KCal::Person person(att.name, att.email);
			event->setOrganizer(person);
		} else {
			KCal::Attendee *person = new KCal::Attendee(att.name, att.email);
			event->addAttendee(person);
		}
	}

	if (message.recurrency.isRecurring()) {
		// if this event is a recurring event create the recurrency
		createKCalRecurrency(event->recurrence(), message.recurrency);
	}

	// TODO add further message

	item.setPayload( KCal::Incidence::Ptr(event) );

	item.setModificationTime( message.modified );

	// notify akonadi about the new data
	itemRetrieved(item);
	return true;
}

void ExCalResource::aboutToQuit()
{
  // TODO: any cleanup you need to do while there is still an active
  // event loop. The resource will terminate after this method returns
}

void ExCalResource::configure( WId windowId )
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

void ExCalResource::itemAdded( const Akonadi::Item &item, const Akonadi::Collection &collection )
{
  Q_UNUSED( item );
  Q_UNUSED( collection );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has created an item in a collection managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

/**
 * Called when somebody else, e.g. a client application, has changed an item 
 * managed this resource.
 */
void ExCalResource::itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts)
{
        Q_UNUSED(parts);
	return;

        // Get the payload for the item.
	kWarning() << "fetch cached item: {" <<
		currentCollection().name() << "," << item.id() << "} = {" <<
		currentCollection().remoteId() << "," << item.remoteId() << "}";
        Akonadi::ItemFetchJob *job = new Akonadi::ItemFetchJob(item);
        connect(job, SIGNAL(result(KJob*)), SLOT(itemChangedContinue(KJob*)));
        job->fetchScope().fetchFullPayload();
}

/**
 * Finish changing an item, now that we (hopefully!) have its payload in hand.
 */
void ExCalResource::itemChangedContinue(KJob* job)
{
        if (job->error()) {
            emit status(Broken, i18n("Failed to get cached data"));
            return;
        }
        Akonadi::ItemFetchJob *fetchJob = qobject_cast<Akonadi::ItemFetchJob*>(job);
        const Akonadi::Item item = fetchJob->items().first();

	if (!logon()) {
		return;
	}

	qulonglong messageId = item.remoteId().toULongLong();
	qulonglong folderId = currentCollection().remoteId().toULongLong();
	MapiAppointment message(m_connection, "ExCalResource::itemChangedContinue", messageId);
	if (!message.open(folderId)) {
		kError() << "open failed!";
		emit status(Broken, i18n("Unable to open item: { %1, %2 }", currentCollection().name(), messageId));
		return;
	}

	// find the remoteId of the item and the collection and try to fetch the needed data from the server
	kWarning() << "fetching item: {" << 
		currentCollection().name() << "," << item.id() << "} = {" << 
		folderId << "," << messageId << "}";
	emit status(Running, i18n("Fetching item: { %1, %2 }", currentCollection().name(), messageId));
	if (!message.propertiesPull()) {
		kError() << "propertiesPull failed!";
		emit status(Broken, i18n("Unable to fetch item: { %1, %2 }", currentCollection().name(), messageId));
		return;
	}
        kWarning() << "got item data:" << message.id << ":" << message.title;

	// Extract the event from the item.
	KCal::Event::Ptr event = item.payload<KCal::Event::Ptr>();
	Q_ASSERT(event->setUid == item.remoteId());
	message.title = event->summary();
	message.begin.setTime_t(event->dtStart().toTime_t());
	message.end.setTime_t(event->dtEnd().toTime_t());
	Q_ASSERT(message.created == event->created());

	// Check that between the item being modified, and this update attempt
	// that no conflicting update happened on the server.
	if (event->lastModified() < KDateTime(message.modified)) {
		kWarning() << "Exchange data modified more recently" << event->lastModified()
			<< "than cached data" << KDateTime(message.modified);
		// TBD: Update cache with data from Exchange.
		return;
	}
	message.modified = item.modificationTime();
	message.text = event->description();
	message.sender = event->organizer().name();
	message.location = event->location();
	if (event->alarms().count()) {
		KCal::Alarm* alarm = event->alarms().first();
		message.reminderActive = true;
		// TODO Maybe we should check which one is set and then use either the time or the delte
		// KDateTime reminder(message.reminderTime);
		// reminder.setTimeSpec( KDateTime::Spec(KDateTime::UTC) );
		// alarm->setTime( reminder );
		message.reminderMinutes = alarm->startOffset() / -60;
	} else {
		message.reminderActive = false;
	}

	message.attendees.clear();
	Attendee att;
	att.name = event->organizer().name();
	att.email = event->organizer().email();
	att.setOrganizer(true);
	message.attendees.append(att);
	att.setOrganizer(false);
	foreach (KCal::Attendee *person, event->attendees()) {
	att.name = person->name();
	att.email = person->email();
	message.attendees.append(att);
	}

	if (message.recurrency.isRecurring()) {
		// if this event is a recurring event create the recurrency
//                createKCalRecurrency(event->recurrence(), message.recurrency);
	}

	// TODO add further data

	// Update exchange with the new message.
	kWarning() << "updating item: {" << 
		currentCollection().name() << "," << item.id() << "} = {" << 
		folderId << "," << messageId << "}";
	emit status(Running, i18n("Updating item: { %1, %2 }", currentCollection().name(), messageId));
	if (!message.propertiesPush()) {
		kError() << "propertiesPush failed!";
		emit status(Running, i18n("Failed to update: { %1, %2 }", currentCollection().name(), messageId));
		return;
	}
	changeCommitted(item);
}

void ExCalResource::itemRemoved( const Akonadi::Item &item )
{
  Q_UNUSED( item );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has deleted an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}


void ExCalResource::createKCalRecurrency(KCal::Recurrence* rec, const MapiRecurrencyPattern& pattern)
{
	rec->clear();

	switch (pattern.mRecurrencyType) {
		case MapiRecurrencyPattern::Daily:
			rec->setDaily(pattern.mPeriod);
			break;
		case MapiRecurrencyPattern::Weekly:
		case MapiRecurrencyPattern::Every_Weekday:
			rec->setWeekly(pattern.mPeriod, pattern.mDays, pattern.mFirstDOW);
			break;
		case MapiRecurrencyPattern::Monthly:
			rec->setMonthly(pattern.mPeriod);
			break;
		case MapiRecurrencyPattern::Yearly:
			rec->setYearly(pattern.mPeriod);
			break;
		default:
			kDebug() << "uncaught recurrency type:"<<pattern.mRecurrencyType;
			return;
		// TODO there are further recurrency types in exchange like e.g. 1st every month, ...
	}

	if (pattern.mEndType == MapiRecurrencyPattern::Count) {
		rec->setDuration(pattern.mOccurrenceCount);
	} else if (pattern.mEndType == MapiRecurrencyPattern::Date) {
		rec->setEndDateTime(KDateTime(pattern.mEndDate));
	} 
}

bool ExCalResource::logon(void)
{
	if (!m_connected) {
		// logon to exchange (if needed)
		emit status(Running, i18n("Logging in to Exchange"));
		m_connected = m_connection->login(Settings::self()->profileName());
	}
	if (!m_connected) {
		emit status(Broken, i18n("Unable to login as %1").arg(Settings::self()->profileName()));
	}
	return m_connected;
}

void ExCalResource::logoff(void)
{
	// There is no logoff operation. We just want to make sure we retry the
	// logon next time.
	m_connected = false;
}

AKONADI_RESOURCE_MAIN( ExCalResource )

#include "excalresource.moc"
