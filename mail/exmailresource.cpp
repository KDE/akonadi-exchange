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

#include "exmailresource.h"

#include <QtDBus/QDBusConnection>

#include <KLocalizedString>
#include <KWindowSystem>
#include <KStandardDirs>

#include <Akonadi/AgentManager>
#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>
#include <akonadi/kmime/messageparts.h>
#include <kmime/kmime_message.h>

#include "mapiconnector2.h"
#include "profiledialog.h"

using namespace Akonadi;

ExMailResource::ExMailResource(const QString &id) :
	MapiResource(id, i18n("Exchange Mail"), IPF_NOTE, "IPM.Note", KMime::Message::mimeType())
{
	new SettingsAdaptor(Settings::self());
	QDBusConnection::sessionBus().registerObject(QLatin1String("/Settings"),
						     Settings::self(),
						     QDBusConnection::ExportAdaptors);
}

ExMailResource::~ExMailResource()
{
}

void ExMailResource::retrieveCollections()
{
	// First select who to log in as.
	profileSet(Settings::self()->profileName());

	Collection::List collections;
	fetchCollections(TopInformationStore, collections);

	// Notify Akonadi about the new collections.
	collectionsRetrieved(collections);
}

void ExMailResource::retrieveItems(const Akonadi::Collection &collection)
{
	Item::List items;
	Item::List deletedItems;
	
	fetchItems(collection, items, deletedItems);
	kError() << "new/changed items:" << items.size() << "deleted items:" << deletedItems.size();
	itemsRetrievedIncremental(items, deletedItems);
}

bool ExMailResource::retrieveItem(const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts)
{
	Q_UNUSED(parts);

	MapiNote *message = fetchItem<MapiNote>(itemOrig);
	if (!message) {
		return false;
	}

	// Create a clone of the passed in const Item and fill it with the payload.
	Akonadi::Item item(itemOrig);
/*
	item.setMimeType(KMime::Message::mimeType());
	item.setPayload(KMime::Message::Ptr(message));

	// update status flags
	if (KMime::isSigned(message.get())) {
		item.setFlag(Akonadi::MessageFlags::Signed);
	}
	if (KMime::isEncrypted(message.get())) {
		item.setFlag(Akonadi::MessageFlags::Encrypted);
	}
	if (KMime::isInvitation(message.get())) {
		item.setFlag(Akonadi::MessageFlags::HasInvitation);
	}
	if (KMime::hasAttachment(message.get())) {
		item.setFlag(Akonadi::MessageFlags::HasAttachment);
	}
*/
	// TODO add further message properties.
//	item.setModificationTime(message->modified);
	delete message;

	// Notify Akonadi about the new data.
	itemRetrieved(item);
	return true;
}

void ExMailResource::aboutToQuit()
{
  // TODO: any cleanup you need to do while there is still an active
  // event loop. The resource will terminate after this method returns
}

void ExMailResource::configure( WId windowId )
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

void ExMailResource::itemAdded( const Akonadi::Item &item, const Akonadi::Collection &collection )
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
void ExMailResource::itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts)
{
        Q_UNUSED(parts);

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
void ExMailResource::itemChangedContinue(KJob* job)
{
        if (job->error()) {
            emit status(Broken, i18n("Failed to get cached data"));
            return;
        }
        Akonadi::ItemFetchJob *fetchJob = qobject_cast<Akonadi::ItemFetchJob*>(job);
        const Akonadi::Item item = fetchJob->items().first();

	MapiNote *message = fetchItem<MapiNote>(item);
	if (!message) {
		return;
	}

	// Extract the event from the item.
#if 0
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
#endif
}

void ExMailResource::itemRemoved( const Akonadi::Item &item )
{
  Q_UNUSED( item );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has deleted an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

AKONADI_RESOURCE_MAIN( ExMailResource )

#include "exmailresource.moc"
