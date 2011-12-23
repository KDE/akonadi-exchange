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

#include "settings.h"
#include "settingsadaptor.h"

#include "excalresource.h"

#include <QtDBus/QDBusConnection>

#include <KLocalizedString>
#include <KWindowSystem>

#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>

#include <KCal/Event>

#include "mapiconnector2.h"
#include "profiledialog.h"

/**
 * Display ids in the same format we use when stored in Akonadi.
 */
#define ID_BASE 36

/**
 * Set this to 1 to pull all the properties, e.g. to see what a server has
 * available.
 */
#ifndef DEBUG_APPOINTMENT_PROPERTIES
#define DEBUG_APPOINTMENT_PROPERTIES 0
#endif

using namespace Akonadi;

/**
 * An Appointment, with attendee recipients.
 *
 * The attendees include not just the To/CC/BCC recipients (@ref propertiesPull)
 * but also whoever the meeting was sent-on-behalf-of. That might or might 
 * not be the sender of the invitation.
 */
class MapiAppointment : public MapiMessage
{
public:
	MapiAppointment(MapiConnector2 *connection, const char *tallocName, mapi_id_t folderId, mapi_id_t id);

	/**
	 * Fetch all properties.
	 */
	virtual bool propertiesPull();

	/**
	 * Update a calendar item.
	 */
	virtual bool propertiesPush();

	QString title;
	QString text;
	QString location;
	QDateTime created;
	QDateTime begin;
	QDateTime end;
	QDateTime modified;
	bool reminderActive;
	QDateTime reminderTime;
	uint32_t reminderMinutes;

	MapiRecurrencyPattern recurrency;

private:
	bool debugRecurrencyPattern(RecurrencePattern *pattern);

	/**
	 * Fetch calendar properties.
	 */
	virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended);

	virtual QDebug debug() const;
	virtual QDebug error() const;
};

ExCalResource::ExCalResource(const QString &id) :
	MapiResource(id, i18n("Exchange Calendar"), IPF_APPOINTMENT, "IPM.Appointment", QString::fromAscii("text/calendar"))
{
	new SettingsAdaptor(Settings::self());
	QDBusConnection::sessionBus().registerObject(QLatin1String("/Settings"),
						     Settings::self(),
						     QDBusConnection::ExportAdaptors);
}

ExCalResource::~ExCalResource()
{
}

void ExCalResource::retrieveCollections()
{
	// First select who to log in as.
	profileSet(Settings::self()->profileName());

	Collection::List collections;
	fetchCollections(Calendar, collections);

	// Notify Akonadi about the new collections.
	collectionsRetrieved(collections);
}

void ExCalResource::retrieveItems(const Akonadi::Collection &collection)
{
	Item::List items;
	Item::List deletedItems;
	
	fetchItems(collection, items, deletedItems);
	kError() << "new/changed items:" << items.size() << "deleted items:" << deletedItems.size();
	itemsRetrievedIncremental(items, deletedItems);
}

bool ExCalResource::retrieveItem( const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts )
{
	Q_UNUSED(parts);

	MapiAppointment *message = fetchItem<MapiAppointment>(itemOrig);
	if (!message) {
		return false;
	}

	// Create a clone of the passed in Item and fill it with the payload.
	Akonadi::Item item(itemOrig);

	KCal::Event* event = new KCal::Event;
	event->setUid(item.remoteId());
	event->setSummary(message->title);
	event->setDtStart(KDateTime(message->begin));
	event->setDtEnd(KDateTime(message->end));
	event->setCreated(KDateTime(message->created));
	event->setLastModified(KDateTime(message->modified));
	event->setDescription(message->text);
	foreach (MapiRecipient recipient, message->recipients()) {
		if (recipient.type() == MapiRecipient::Sender) {
			KCal::Person person(recipient.name, recipient.email);
			event->setOrganizer(person);
		} else {
			KCal::Attendee *person = new KCal::Attendee(recipient.name, recipient.email);
			event->addAttendee(person);
		}
	}

	event->setLocation(message->location);
	if (message->reminderActive) {
		KCal::Alarm* alarm = new KCal::Alarm(dynamic_cast<KCal::Incidence*>(event));
		// TODO Maybe we should check which one is set and then use either the time or the delte
		// KDateTime reminder(message->reminderTime);
		// reminder.setTimeSpec(KDateTime::Spec(KDateTime::UTC));
		// alarm->setTime(reminder);
		alarm->setStartOffset(KCal::Duration(message->reminderMinutes * -60));
		alarm->setEnabled(true);
		event->addAlarm(alarm);
	}

	if (message->recurrency.isRecurring()) {
		// if this event is a recurring event create the recurrency
		createKCalRecurrency(event->recurrence(), message->recurrency);
	}
	item.setPayload(KCal::Incidence::Ptr(event));

	// TODO add further message properties.
	item.setModificationTime(message->modified);
	delete message;

	// Notify Akonadi about the new data.
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

	MapiAppointment *message = fetchItem<MapiAppointment>(item);
	if (!message) {
		return;
	}

	// Extract the event from the item.
	KCal::Event::Ptr event = item.payload<KCal::Event::Ptr>();
	Q_ASSERT(event->setUid == item.remoteId());
	message->title = event->summary();
	message->begin.setTime_t(event->dtStart().toTime_t());
	message->end.setTime_t(event->dtEnd().toTime_t());
	Q_ASSERT(message->created == event->created());

	// Check that between the item being modified, and this update attempt
	// that no conflicting update happened on the server.
	if (event->lastModified() < KDateTime(message->modified)) {
		kWarning() << "Exchange data modified more recently" << event->lastModified()
			<< "than cached data" << KDateTime(message->modified);
		// TBD: Update cache with data from Exchange.
		return;
	}
	message->modified = item.modificationTime();
	message->text = event->description();
	//message->sender = event->organizer().name();
	message->location = event->location();
	if (event->alarms().count()) {
		KCal::Alarm* alarm = event->alarms().first();
		message->reminderActive = true;
		// TODO Maybe we should check which one is set and then use either the time or the delte
		// KDateTime reminder(message->reminderTime);
		// reminder.setTimeSpec( KDateTime::Spec(KDateTime::UTC) );
		// alarm->setTime( reminder );
		message->reminderMinutes = alarm->startOffset() / -60;
	} else {
		message->reminderActive = false;
	}

	MapiRecipient att(MapiRecipient::Sender);
	att.name = event->organizer().name();
	att.email = event->organizer().email();
	message->addUniqueRecipient("event organiser", att);

	att.setType(MapiRecipient::To);
	foreach (KCal::Attendee *person, event->attendees()) {
		att.name = person->name();
		att.email = person->email();
		message->addUniqueRecipient("event attendee", att);
	}

	if (message->recurrency.isRecurring()) {
		// if this event is a recurring event create the recurrency
//                createKCalRecurrency(event->recurrence(), message->recurrency);
	}

	// TODO add further data

	// Update exchange with the new message->
	kWarning() << "updating item: {" << 
		currentCollection().name() << "," << item.id() << "} = {" << 
		currentCollection().remoteId() << "," << item.remoteId() << "}";
	emit status(Running, i18n("Updating item: { %1, %2 }", currentCollection().name(), item.id()));
	if (!message->propertiesPush()) {
		error(*message, i18n("unable to update item"));
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

MapiAppointment::MapiAppointment(MapiConnector2 *connector, const char *tallocName, mapi_id_t folderId, mapi_id_t id) :
	MapiMessage(connector, tallocName, folderId, id)
{
}

QDebug MapiAppointment::debug() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE)) /*<< title*/;
}

QDebug MapiAppointment::error() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1/%2:");
	return MapiObject::error(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE)) /*<< title*/;
}

bool MapiAppointment::debugRecurrencyPattern(RecurrencePattern *pattern)
{
	// do the actual work
	debug() << "-- Recurrency debug output [BEGIN] --";
	switch (pattern->RecurFrequency) {
	case RecurFrequency_Daily:
		debug() << "Fequency: daily";
		break;
	case RecurFrequency_Weekly:
		debug() << "Fequency: weekly";
		break;
	case RecurFrequency_Monthly:
		debug() << "Fequency: monthly";
		break;
	case RecurFrequency_Yearly:
		debug() << "Fequency: yearly";
		break;
	default:
		debug() << "unsupported frequency:"<<pattern->RecurFrequency;
		return false;
	}

	switch (pattern->PatternType) {
	case PatternType_Day:
		debug() << "PatternType: day";
		break;
	case PatternType_Week:
		debug() << "PatternType: week";
		break;
	case PatternType_Month:
		debug() << "PatternType: month";
		break;
	default:
		debug() << "unsupported patterntype:"<<pattern->PatternType;
		return false;
	}

	debug() << "Calendar:" << pattern->CalendarType;
	debug() << "FirstDateTime:" << pattern->FirstDateTime;
	debug() << "Period:" << pattern->Period;
	if (pattern->PatternType == PatternType_Month) {
		debug() << "PatternTypeSpecific:" << pattern->PatternTypeSpecific.Day;
	} else if (pattern->PatternType == PatternType_Week) {
		debug() << "PatternTypeSpecific:" << QString::number(pattern->PatternTypeSpecific.WeekRecurrencePattern, 2);
	}

	switch (pattern->EndType) {
		case END_AFTER_DATE:
			debug() << "EndType: after date";
			break;
		case END_AFTER_N_OCCURRENCES:
			debug() << "EndType: after occurenc count";
			break;
		case END_NEVER_END:
			debug() << "EndType: never";
			break;
		default:
			debug() << "unsupported endtype:"<<pattern->EndType;
			return false;
	}
	debug() << "OccurencCount:" << pattern->OccurrenceCount;
	debug() << "FirstDOW:" << pattern->FirstDOW;
	debug() << "Start:" << pattern->StartDate;
	debug() << "End:" << pattern->EndDate;
	debug() << "-- Recurrency debug output [END] --";

	return true;
}

bool MapiAppointment::propertiesPull(QVector<int> &tags, const bool tagsAppended)
{
	/**
	 * The list of tags used to fetch an Appointment.
	 */
	static int ourTagList[] = {
		PidTagMessageClass,
		PidTagDisplayTo,
		PidTagConversationTopic,
		PidTagBody,
		PidTagLastModificationTime,
		PidTagCreationTime,
		PidTagStartDate,
		PidTagEndDate,
		PidLidLocation,
		PidLidReminderSet,
		PidLidReminderSignalTime,
		PidLidReminderDelta,
		PidLidRecurrenceType,
		PidLidAppointmentRecur,

		PidTagSentRepresentingEmailAddress,
		PidTagSentRepresentingEmailAddress_string8,

		PidTagSentRepresentingName,
		PidTagSentRepresentingName_string8,

		PidTagSentRepresentingSimpleDisplayName,
		PidTagSentRepresentingSimpleDisplayName_string8,

		PidTagOriginalSentRepresentingEmailAddress,
		PidTagOriginalSentRepresentingEmailAddress_string8,

		PidTagOriginalSentRepresentingName,
		PidTagOriginalSentRepresentingName_string8,
		0 };
	static SPropTagArray ourTags = {
		(sizeof(ourTagList) / sizeof(ourTagList[0])) - 1,
		(MAPITAGS *)ourTagList };

	if (!tagsAppended) {
		for (unsigned i = 0; i < ourTags.cValues; i++) {
			int newTag = ourTags.aulPropTag[i];
			
			if (!tags.contains(newTag)) {
				tags.append(newTag);
			}
		}
	}
	if (!MapiMessage::propertiesPull(tags, tagsAppended)) {
		return false;
	}

	// Start with a clean slate.
	reminderActive = false;

	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	unsigned recurrenceType = 0;
	RecurrencePattern *pattern = 0;

	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		switch (property.tag()) {
		case PidTagMessageClass:
			// Sanity check the message class.
			if (QLatin1String("IPM.Appointment") != property.value().toString()) {
				// this one is not an appointment
				return false;
			}
			break;
		case PidTagConversationTopic:
			title = property.value().toString();
			break;
		case PidTagBody:
			text = property.value().toString();
			break;
		case PidTagLastModificationTime:
			modified = property.value().toDateTime();
			break;
		case PidTagCreationTime:
			created = property.value().toDateTime();
			break;
		case PidTagStartDate:
			begin = property.value().toDateTime();
			break;
		case PidTagEndDate:
			end = property.value().toDateTime();
			break;
		case PidLidLocation:
			location = property.value().toString();
			break;
		case PidLidReminderSet:
			reminderActive = property.value().toInt();
			break;
		case PidLidReminderSignalTime:
			reminderTime = property.value().toDateTime();
			break;
		case PidLidReminderDelta:
			reminderMinutes = property.value().toInt();
			break;
		case PidLidRecurrenceType:
			recurrenceType = property.value().toInt();
			break;
		case PidLidAppointmentRecur:
			pattern = get_RecurrencePattern(ctx(), &m_properties[i].value.bin);
			break;
		default:
#if (DEBUG_APPOINTMENT_PROPERTIES)
			debug() << "ignoring appointment property:" << tagName(property.tag()) << property.value();
#endif
			break;
		}
	}

	// Is there a recurrance type set?
	if (recurrenceType != 0) {
		if (pattern) {
			debugRecurrencyPattern(pattern);
			recurrency.setData(pattern);
		} else {
			// TODO This should not happen. PidLidRecurrenceType says this is a recurring event, so why is there no PidLidAppointmentRecur???
			debug() << "missing recurrencePattern";
			return false;
		}
	}
	return true;
}

bool MapiAppointment::propertiesPull()
{
	static bool tagsAppended = false;
	static QVector<int> tags;

	if (!propertiesPull(tags, tagsAppended)) {
		tagsAppended = true;
		return false;
	}
	tagsAppended = true;
	return true;
}

bool MapiAppointment::propertiesPush()
{
	// Overwrite all the fields we know about.
	if (!propertyWrite(PidTagConversationTopic, title)) {
		return false;
	}
	if (!propertyWrite(PidTagBody, text)) {
		return false;
	}
	if (!propertyWrite(PidTagLastModificationTime, modified)) {
		return false;
	}
	if (!propertyWrite(PidTagCreationTime, created)) {
		return false;
	}
	if (!propertyWrite(PidTagStartDate, begin)) {
		return false;
	}
	if (!propertyWrite(PidTagEndDate, end)) {
		return false;
	}
	//if (!propertyWrite(PidTagSentRepresentingName, sender()[0].name)) {
	//	return false;
	//}
	if (!propertyWrite(PidLidLocation, location)) {
		return false;
	}
	if (!propertyWrite(PidLidReminderSet, reminderActive)) {
		return false;
	}
	if (reminderActive) {
		if (!propertyWrite(PidLidReminderSignalTime, reminderTime)) {
			return false;
		}
		if (!propertyWrite(PidLidReminderDelta, reminderMinutes)) {
			return false;
		}
	}
#if 0
	const char* toAttendeesString = (const char *)find_mapi_SPropValue_data(&properties_array, PR_DISPLAY_TO_UNICODE);
	uint32_t* recurrenceType = (uint32_t*)find_mapi_SPropValue_data(&properties_array, PidLidRecurrenceType);
	Binary_r* binDataRecurrency = (Binary_r*)find_mapi_SPropValue_data(&properties_array, PidLidAppointmentRecur);

	if (recurrenceType && (*recurrenceType) > 0x0) {
		if (binDataRecurrency != 0x0) {
			RecurrencePattern* pattern = get_RecurrencePattern(mem_ctx, binDataRecurrency);
			debugRecurrencyPattern(pattern);

			data.recurrency.setData(pattern);
		} else {
			// TODO This should not happen. PidLidRecurrenceType says this is a recurring event, so why is there no PidLidAppointmentRecur???
			debug() << "missing recurrencePattern in message"<<messageID<<"in folder"<<folderID;
			}
	}

	getAttendees(obj_message, QString::fromLocal8Bit(toAttendeesString), data);
#endif
	if (!MapiMessage::propertiesPush()) {
		return false;
	}
#if 0
	debug() << "************  OpenFolder";
	if (!OpenFolder(&m_store, folder.id(), folder.d())) {
		error() << "cannot open folder" << folderID
			<< ", error:" << mapiError();
		return false;
        }
	debug() << "************  SaveChangesMessage";
	if (!SaveChangesMessage(folder.d(), message.d(), KeepOpenReadWrite)) {
		error() << "cannot save message" << messageID << "in folder" << folderID
			<< ", error:" << mapiError();
		return false;
	}
#endif
	debug() << "************  SubmitMessage";
	if (MAPI_E_SUCCESS != SubmitMessage(&m_object)) {
		error() << "cannot submit message, error:" << mapiError();
		return false;
	}
	struct mapi_SPropValue_array replyProperties;
	debug() << "************  TransportSend";
	if (MAPI_E_SUCCESS != TransportSend(&m_object, &replyProperties)) {
		error() << "cannot send message, error:" << mapiError();
		return false;
	}
	return true;
}

AKONADI_RESOURCE_MAIN( ExCalResource )

#include "excalresource.moc"
