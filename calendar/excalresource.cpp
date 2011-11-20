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
  : ResourceBase( id )
{
	new SettingsAdaptor( Settings::self() );
	QDBusConnection::sessionBus().registerObject( QLatin1String( "/Settings" ),
							Settings::self(), QDBusConnection::ExportAdaptors );

	// initialize the mapi connector
	connector = new MapiConnector2;
}

ExCalResource::~ExCalResource()
{
	if (connector)
		delete connector;
}

void ExCalResource::retrieveCollections()
{
	kDebug() << "retrieveCollections() called";

	emit status(Running, i18n("Logging in to Exchange"));
	bool ok = connector->login(Settings::self()->profileName());
	if (!ok) {
		emit status(Broken, i18n("Unable to login") );
		return;
	}

	Collection::List collections;

	QStringList folderMimeType;
	folderMimeType << QString::fromAscii("text/calendar");

	// create the new collection
	Collection root;
	root.setParentCollection(Collection::root());
	root.setContentMimeTypes(folderMimeType);

	QList<FolderData> list;
	emit status(Running, i18n("Fetching folder list from Exchange"));
	if (connector->fetchFolderList(list, 0x0, QString::fromAscii(IPF_APPOINTMENT))) {
		foreach (FolderData data, list) {
			// TODO: take the first calender for now, but Exchange might have more calendar folders
			root.setRemoteId(data.id);
			root.setName(i18n("Exchange: ") + data.name);
			break;
		}
	}
	collections.append(root);

	// notify akonadi about the new calendar collection
	collectionsRetrieved(collections);
}

void ExCalResource::retrieveItems( const Akonadi::Collection &collection )
{
	kDebug() << "retrieveItems() called for collection "<< collection.id();

	// logon to exchange (if needed)
	emit status(Running, i18n("Logging in to Exchange"));
	bool ok = connector->login(Settings::self()->profileName());
	if (!ok) {
		emit status(Broken, i18n("Unable to login") );
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

	// get the folder content for the collection
	Item::List items;
	QList<CalendarDataShort> list;
	emit status(Running, i18n("Fetching calendar data from Exchange"));
	if (connector->fetchFolderContent(collection.remoteId().toULongLong(), list)) {
		kDebug() << "size of collection" << collection.id() << "is" << list.size();

		QSet<QString> checkedRemoteIds;
		// run though all the found data...
		foreach (const CalendarDataShort data, list) {

			checkedRemoteIds << data.id; // store for later use

			if (!knownRemoteIds.contains(data.id)) {
				// we do not know this remoteID -> create a new empty item for it
				Item item(QString::fromAscii("text/calendar"));
				item.setParentCollection(collection);
				item.setRemoteId(data.id);
				item.setRemoteRevision(QString::number(1));
				items << item;
			} else {
				// this item is already known, check if it was update in the meanwhile
				Item& existingItem = knownItems.find(data.id).value();
// 				kDebug() << "Item("<<existingItem.id()<<":"<<data.id<<":"<<existingItem.revision()<<") is already known [Cache-ModTime:"<<existingItem.modificationTime()
// 						<<" Server-ModTime:"<<data.modified<<"] Flags:"<<existingItem.flags()<<"Attrib:"<<existingItem.attributes();
				if (existingItem.modificationTime() < data.modified) {
					kDebug() << existingItem.id()<<"=> this item has changed";

					// force akonadi to call retrieveItem() for this item in order to get updated data
					int revision = existingItem.remoteRevision().toInt();
					existingItem.clearPayload();
					existingItem.setRemoteRevision( QString::number(++revision) );
					items << existingItem;
				}
			}

			// TODO just for debugging...
// 			if (items.size() > 3) {
// 				break;
// 			}
		}

		// now check if some of the items need to be removed
		knownRemoteIds.subtract(checkedRemoteIds);

		Item::List deletedItems;
		foreach (const QString remoteId, knownRemoteIds) {
			deletedItems << knownItems.value(remoteId);
		}

		kDebug() << "itemsRetrievedIncremental(): fetched"<<items.size()<<"new/changed items; delete"<<deletedItems.size()<<"items";
		itemsRetrievedIncremental(items, deletedItems);
	}

	foreach(Item item, items) {
		kDebug() << "[Item-Dump] ID:"<<item.id()<<"RemoteId:"<<item.remoteId()<<"Revision:"<<item.revision()<<"ModTime:"<<item.modificationTime();
	}
}

bool ExCalResource::retrieveItem( const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts )
{
	Q_UNUSED( parts );

	kDebug() << "retrieveItem() called for item "<< itemOrig.id() << "remoteId:" << itemOrig.remoteId();

	// logon to exchange (if needed)
	emit status(Running, i18n("Logging in to Exchange"));
	connector->login(Settings::self()->profileName());

	// find the remoteId of the item and the collection and try to fetch the needed data from the server
	CalendarData data;
	kDebug() << "currentColleaction:"<<currentCollection().id()<<":"<<currentCollection().remoteId();
	kDebug() << "fetching data for item:"<<itemOrig.id()<<":"<<itemOrig.remoteId().toULongLong();
	emit status(Running, i18n("Fetching appointment %1", itemOrig.remoteId()));
	if (connector->fetchCalendarData(currentCollection().remoteId().toULongLong(), itemOrig.remoteId().toULongLong(), data)) {
		kDebug() << "got data; item:"<<data.id<<":"<<data.title;

		// Create a clone of the passed in Item and fill it with the payload
		Akonadi::Item item(itemOrig);

		KCal::Event* event = new KCal::Event;
		event->setUid(item.remoteId());
		event->setSummary( data.title );
		event->setDtStart( KDateTime(data.begin) );
		event->setDtEnd( KDateTime(data.end) );
		event->setCreated( KDateTime(data.created) );
		event->setLastModified( KDateTime(data.modified) );
		event->setDescription( data.text );
		event->setOrganizer( data.sender );
		event->setLocation( data.location );
		if (data.reminderActive) {
			KCal::Alarm* alarm = new KCal::Alarm( dynamic_cast<KCal::Incidence*>( event ) );
			// TODO Maybe we should check which one is set and then use either the time or the delte
			// KDateTime reminder(data.reminderTime);
			// reminder.setTimeSpec( KDateTime::Spec(KDateTime::UTC) );
			// alarm->setTime( reminder );
			alarm->setStartOffset(KCal::Duration(data.reminderDelta*-60));
			alarm->setEnabled( true );
			event->addAlarm( alarm );
		}

		foreach (Attendee att, data.anttendees) {
			if (att.isOranizer) {
				KCal::Person person(att.name, att.email);
				event->setOrganizer(person);
			} else {
				KCal::Attendee *person = new KCal::Attendee(att.name, att.email);
				event->addAttendee(person);
			}
		}

		if (data.recurrency.isRecurring()) {
			// if this event is a recurring event create the recurrency
			createKCalRecurrency(event->recurrence(), data.recurrency);
		}

		// TODO add further data

		item.setPayload( KCal::Incidence::Ptr(event) );

		item.setModificationTime( data.modified );

		// notify akonadi about the new data
		itemRetrieved(item);

		return true;
	}

	kDebug() << "Failed to get data for item"<<itemOrig.remoteId();
	return false;
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

void ExCalResource::itemChanged( const Akonadi::Item &item, const QSet<QByteArray> &parts )
{
  Q_UNUSED( item );
  Q_UNUSED( parts );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has changed an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
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


AKONADI_RESOURCE_MAIN( ExCalResource )

#include "excalresource.moc"
