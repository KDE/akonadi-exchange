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

#include <Akonadi/CachePolicy>
#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>

#include <KCal/Event>

#include "mapiconnector2.h"
#include "profiledialog.h"

/**
 * Set this to 1 to pull all the properties, e.g. to see what a server has
 * available.
 */
#ifndef DEBUG_APPOINTMENT_PROPERTIES
#define DEBUG_APPOINTMENT_PROPERTIES 0
#endif

using namespace Akonadi;

static QBitArray ex2kcalRecurrenceDays(uint32_t exchangeDays)
{
	QBitArray bitArray(7, false);

	if (exchangeDays & 0x00000001) // Sunday
		bitArray.setBit(6, true);
	if (exchangeDays & 0x00000002) // Monday
		bitArray.setBit(0, true);
	if (exchangeDays & 0x00000004) // Tuesday
		bitArray.setBit(1, true);
	if (exchangeDays & 0x00000008) // Wednesday
		bitArray.setBit(2, true);
	if (exchangeDays & 0x00000010) // Thursday
		bitArray.setBit(3, true);
	if (exchangeDays & 0x00000020) // Friday
		bitArray.setBit(4, true);
	if (exchangeDays & 0x00000040) // Saturday
		bitArray.setBit(5, true);
	return bitArray;
}

static uint32_t ex2kcalDayOfWeek(uint32_t exchangeDayOfWeek)
{
	uint32_t retVal = exchangeDayOfWeek;
	if (retVal == 0) {
		// Exchange-Sunday(0) mapped to KCal-Sunday(7)
		retVal = 7;
	}
	return retVal;
}

static uint32_t ex2kcalDays(uint32_t exchangeMinutes)
{
	return exchangeMinutes / 60 / 24;
}

static KDateTime ex2kcalTimes(uint32_t exchangeMinutes)
{
	// exchange stores the recurrency times as minutes since 1.1.1601
	QDateTime calc(QDate(1601, 1, 1));
	int days = ex2kcalDays(exchangeMinutes);
	int secs = exchangeMinutes - (days * 24 * 60);
	return KDateTime(calc.addDays(days).addSecs(secs));
}

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
	MapiAppointment(MapiConnector2 *connection, const char *tallocName, MapiId &id);

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

	void ex2kcalRecurrency(KCal::Recurrence *rec);

private:
	RecurrencePattern *m_pattern;

	/**
	 * Fetch calendar properties.
	 */
	virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);

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

const QString ExCalResource::profile()
{
	// First select who to log in as.
	return Settings::self()->profileName();
}

void ExCalResource::retrieveCollections()
{
	Collection::List collections;

	setName(i18n("Exchange Calendar for %1", profile()));
	fetchCollections(Calendar, collections);
	if (collections.size()) {
		Collection root = collections.first();
		Akonadi::CachePolicy cachePolicy;
		cachePolicy.setInheritFromParent(false);
		cachePolicy.setSyncOnDemand(false);
		cachePolicy.setCacheTimeout(-1);
		cachePolicy.setIntervalCheckTime(5);
		root.setCachePolicy(cachePolicy);
	}

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

	KCal::Event *event = new KCal::Event;
	event->setUid(item.remoteId());
	event->setSummary(message->title);
	event->setDtStart(KDateTime(message->begin));
	event->setDtEnd(KDateTime(message->end));
	event->setCreated(KDateTime(message->created));
	event->setLastModified(KDateTime(message->modified));
	event->setDescription(message->text);
	foreach (MapiRecipient recipient, message->recipients()) {
		if (recipient.type() == MapiRecipient::ReplyTo) {
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

	message->ex2kcalRecurrency(event->recurrence());
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

void ExCalResource::configure(WId windowId)
{
	ProfileDialog dlgConfig;
	dlgConfig.setProfileName(Settings::self()->profileName());
	if (windowId) {
		KWindowSystem::setMainWindow(&dlgConfig, windowId);
	}
	if (dlgConfig.exec() == KDialog::Accepted) {
		Settings::self()->setProfileName(dlgConfig.profileName());
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

	// TODO add further data
	message->ex2kcalRecurrency(event->recurrence());

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

MapiAppointment::MapiAppointment(MapiConnector2 *connector, const char *tallocName, MapiId &id) :
	MapiMessage(connector, tallocName, id),
	m_pattern(0)
{
}

QDebug MapiAppointment::debug() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1:");
	return MapiObject::debug(prefix.arg(m_id.toString()));
}

QDebug MapiAppointment::error() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1:");
	return MapiObject::error(prefix.arg(m_id.toString()));
}

void MapiAppointment::ex2kcalRecurrency(KCal::Recurrence *rec)
{
	rec->clear();
	if (!m_pattern) {
		// No recurrency.
		return;
	}
	debug() << "-- Recurrency debug output [BEGIN] --";
	debug() << "Calendar:" << m_pattern->CalendarType;
	if (m_pattern->RecurFrequency == RecurFrequency_Daily && m_pattern->PatternType == PatternType_Day) {
		debug() << "Every n days:" << ex2kcalDays(m_pattern->Period);
		rec->setDaily(ex2kcalDays(m_pattern->Period));
	}
	else if (m_pattern->RecurFrequency == RecurFrequency_Daily && m_pattern->PatternType == PatternType_Week) {
		debug() << "Every weekday";
		QBitArray bitArray(7, true); // everyday ...
		bitArray.setBit(5, false); // ... except saturday ..
		bitArray.setBit(6, false); // ... except sunday

		rec->setWeekly(
			ex2kcalDays(m_pattern->Period) / 7,
			bitArray,
			ex2kcalDayOfWeek(m_pattern->FirstDOW));
	}
	else if (m_pattern->RecurFrequency == RecurFrequency_Weekly && m_pattern->PatternType == PatternType_Week) {
		debug() << "Every n weeks on one or more particular days of the week:" << m_pattern->Period <<
			"MTWTFSS:" << ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.WeekRecurrencePattern);
		rec->setWeekly(
			m_pattern->Period,
			ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.WeekRecurrencePattern),
			ex2kcalDayOfWeek(m_pattern->FirstDOW));
	}
	else if (m_pattern->RecurFrequency == RecurFrequency_Monthly) {
		rec->setMonthly(m_pattern->Period);
		debug() << "Every n months:" << m_pattern->Period;
		switch (m_pattern->PatternType)
		{
		case PatternType_Month:
		case PatternType_HjMonth:
			debug() << "...on the nth day:" << m_pattern->PatternTypeSpecific.Day;
			rec->addMonthlyDate(m_pattern->PatternTypeSpecific.Day);
			break;
		case PatternType_MonthNth:
		case PatternType_HjMonthNth:
			debug() << "...on a specific day of the week on the nth week:" <<
				m_pattern->PatternTypeSpecific.MonthRecurrencePattern.N << "MTWTFSS:" <<
				ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern);
			rec->addMonthlyPos(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.N, ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern));
			break;
		case PatternType_MonthEnd:
		case PatternType_HjMonthEnd:
			debug() << "...at the end of the month" << m_pattern->PatternTypeSpecific.Day;
			rec->addMonthlyDate(m_pattern->PatternTypeSpecific.Day);
			break;
		default:
			error() << "unsupported monthly frequency with patterntype:" << m_pattern->PatternType;
			break;
		}
	}
	else if (m_pattern->RecurFrequency == RecurFrequency_Yearly) {
		rec->setYearly(1);
		debug() << "Every year";
		switch (m_pattern->PatternType)
		{
		case PatternType_Month:
		case PatternType_HjMonth:
			debug() << "...on the mth day of the nth month:" << m_pattern->PatternTypeSpecific.Day << m_pattern->Period;
			rec->addYearlyMonth(m_pattern->Period);
			rec->addYearlyDate(m_pattern->PatternTypeSpecific.Day);
			break;
		case PatternType_MonthNth:
		case PatternType_HjMonthNth:
			debug() << "...on a specific day of the week on the mth day of the nth month:" <<
				m_pattern->PatternTypeSpecific.MonthRecurrencePattern.N << m_pattern->Period <<
				"MTWTFSS:" << ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern);
			rec->addYearlyMonth(m_pattern->Period);
			rec->addYearlyPos(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.N, ex2kcalRecurrenceDays(m_pattern->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern));
			break;
		default:
			error() << "unsupported yearly frequency with patterntype:" << m_pattern->PatternType;
			break;
		};
	} else {
		// TODO there are further recurrency types in exchange like e.g. 1st every month, ...
		error() << "unsupported combination of frequency:" << m_pattern->RecurFrequency <<
			"and patterntype:" << m_pattern->PatternType;
	}

	rec->setStartDateTime(ex2kcalTimes(m_pattern->StartDate));
	debug() << "Start:" << rec->startDateTime();
	switch (m_pattern->EndType) {
	case END_AFTER_DATE:
		rec->setEndDateTime(ex2kcalTimes(m_pattern->EndDate));
		debug() << "End:" << rec->endDateTime();
		break;
	case END_AFTER_N_OCCURRENCES:
		rec->setDuration(m_pattern->OccurrenceCount);
		debug() << "End: after occurrence count" << rec->duration();
		break;
	case END_NEVER_END:
	case NEVER_END:
		debug() << "End: never";
		break;
	default:
		error() << "unsupported endtype:" << m_pattern->EndType;
		break;
	}
	rec->dump();
	debug() << "-- Recurrency debug output [END] --";
	//talloc_free(m_pattern);
	//m_pattern = 0;
}

bool MapiAppointment::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
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
		PidTagTransportMessageHeaders,
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
	if (!MapiMessage::propertiesPull(tags, tagsAppended, pullAll)) {
		return false;
	}

	// Start with a clean slate.
	reminderActive = false;

	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	unsigned recurrenceType = 0;
	bool embeddedInBody = false;
	QString header;

	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		switch (property.tag()) {
		case PidTagMessageClass:
			// Sanity check the message class.
			if (QLatin1String("IPM.Appointment") != property.value().toString()) {
				if (QLatin1String("IPM.Note") != property.value().toString()) {
					error() << "retrieved item is not an appointment:" << property.value().toString();
					return false;
				} else {
					embeddedInBody = true;
				}
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
			m_pattern = get_RecurrencePattern(ctx(), &m_properties[i].value.bin);
			break;
		case PidTagTransportMessageHeaders:
			header = property.value().toString();
			break;
		default:
			// Handle oversize objects.
			if (MAPI_E_NOT_ENOUGH_MEMORY == property.value().toInt()) {
				switch (property.tag()) {
				default:
					error() << "missing oversize support:" << tagName(property.tag());
					break;
				}

				// Carry on with next property...
				break;
			}
#if (DEBUG_APPOINTMENT_PROPERTIES)
			debug() << "ignoring appointment property:" << tagName(property.tag()) << property.value();
#endif
			break;
		}
	}

	if (embeddedInBody) {
		// Exchange puts half the information in the headers:
		//
		//	Microsoft Mail Internet Headers Version 2.0
		//	BEGIN:VCALENDAR
		//	PRODID:-//K Desktop Environment//NONSGML libkcal 4.3//EN
		//	VERSION:2.0
		//	BEGIN:VEVENT
		//	DTSTAMP:20100420T092856Z
		//	ORGANIZER;CN="xxx yyy (zzz)":MAILTO:zzz@foo.com
		//	X-UID: 0
		//
		// and the rest in the body:
		//
		//	ATTENDEE;CN="aaa bbb (ccc)";RSVP=TRUE;PARTSTAT=NEEDS-ACTION;
		//	...
		//
		// Unbelievable, but true! Anyway, start by fixing up the 
		// header.
		bool lastChWasNl = false;
		int j = 0;
		for (int i = 0; i < header.size(); i++) {
			QChar ch = header.at(i);
			bool chIsNl = false;

			switch (ch.toAscii()) {
			case '\r':
				// Omit CRs. Propagate the NL state of the 
				// previous character.
				chIsNl = lastChWasNl;
				break;
			case '\n':
				// End with the first double NL.
				if (lastChWasNl) {
					goto DONE;
				}
				chIsNl = true;
				// Copy anything else.
				header[j] = ch;
				j++;
				break;
			default:
				// Copy anything else.
				header[j] = ch;
				j++;
				break;
			}
			lastChWasNl = chIsNl;
		}
DONE:
		header.resize(j);
		text = header + text;
		if (text.isEmpty()) {
			error() << "retrieved content is not an appointment";
			return false;
		}
		return true;
	}

	// Is there a recurrence type set?
	if (recurrenceType != 0) {
		if (!m_pattern) {
			// TODO This should not happen. PidLidRecurrenceType says this is a recurring event, so why is there no PidLidAppointmentRecur???
			debug() << "missing recurrencePattern for recurrenceType:" << recurrenceType;
			recurrenceType = 0;
		}
	}
	return true;
}

bool MapiAppointment::propertiesPull()
{
	static bool tagsAppended = false;
	static QVector<int> tags;

	if (!propertiesPull(tags, tagsAppended, (DEBUG_APPOINTMENT_PROPERTIES) != 0)) {
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
