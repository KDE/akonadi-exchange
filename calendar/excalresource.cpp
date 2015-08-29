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
#include <Akonadi/ItemCreateJob>
#include <Akonadi/ItemDeleteJob>
#include <Akonadi/ItemFetchJob>
#include <Akonadi/ItemFetchScope>

#include <KCalCore/Event>
#include <kcalutils/stringify.h>

#include <libmapi/mapi_nameid.h>
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

static QString stringify(QBitArray &days)
{
    QString result;
    if (days[0]) {
        result.append(i18n("Mo"));
    }
    if (days[1]) {
        result.append(i18n("Tu"));
    }
    if (days[2]) {
        result.append(i18n("We"));
    }
    if (days[3]) {
        result.append(i18n("Th"));
    }
    if (days[4]) {
        result.append(i18n("Fr"));
    }
    if (days[5]) {
        result.append(i18n("Sa"));
    }
    if (days[6]) {
        result.append(i18n("Su"));
    }
    return result;
}

static QString stringify(KDateTime dateTime)
{
    return KCalUtils::Stringify::formatDateTime(dateTime);
}

static QBitArray ex2kcalRecurrenceDays(uint32_t exchangeDays)
{
    QBitArray bitArray(7, false);

    if (exchangeDays & Su) // Sunday
        bitArray.setBit(6, true);
    if (exchangeDays & M) // Monday
        bitArray.setBit(0, true);
    if (exchangeDays & Tu) // Tuesday
        bitArray.setBit(1, true);
    if (exchangeDays & W) // Wednesday
        bitArray.setBit(2, true);
    if (exchangeDays & Th) // Thursday
        bitArray.setBit(3, true);
    if (exchangeDays & F) // Friday
        bitArray.setBit(4, true);
    if (exchangeDays & Sa) // Saturday
        bitArray.setBit(5, true);
    return bitArray;
}

static uint32_t ex2kcalDayOfWeek(uint32_t exchangeDayOfWeek)
{
    uint32_t retVal = exchangeDayOfWeek;
    if (retVal == FirstDOW_Sunday) {
        // Exchange-Sunday(0) mapped to KCal-Sunday(7)
        retVal = 7;
    }
    return retVal;
}

static uint32_t ex2kcalDaysFromMinutes(uint32_t exchangeMinutes)
{
    return exchangeMinutes / 60 / 24;
}

static KDateTime ex2kcalTimes(uint32_t exchangeMinutes)
{
    // exchange stores the recurrency times as minutes since 1.1.1601
    QDateTime calc(QDate(1601, 1, 1));
    int days = ex2kcalDaysFromMinutes(exchangeMinutes);
    int secs = (exchangeMinutes - (days * 24 * 60)) * 60;
    return KDateTime(calc.addDays(days).addSecs(secs));
}

/**
 * An Appointment, with attendee recipients.
 *
 * The attendees include not just the To/CC/BCC recipients (@ref propertiesPull)
 * but also whoever the meeting was sent-on-behalf-of. That might or might
 * not be the sender of the invitation.
 */
class MapiAppointment : public MapiMessage, public KCalCore::Event
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

    // PidLidAppointmentStateFlags
    typedef enum {
        Meeting = 0x1,
        Received = 0x2,
        Canceled = 0x4,
    } AppointmentState;
    Q_DECLARE_FLAGS(AppointmentStates, AppointmentState);

    // PidLidResponseStatus
    typedef enum {
        None,
        Organized,
        Tentative,
        Accepted,
        Declined,
        NotResponded
    } ResponseStatus;

    QList<class MapiAppointmentException *> m_exceptions;

protected:
    virtual QDebug debug() const;
    virtual QDebug error() const;

private:
    /**
     * Populate the object with property contents.
     */
    bool preparePayload();

    /**
     * Fetch calendar properties.
     */
    virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);

    void ex2kcalRecurrency(AppointmentRecurrencePattern *pattern, KCalCore::Recurrence *kcal);
};

Q_DECLARE_OPERATORS_FOR_FLAGS(MapiAppointment::AppointmentStates);

class MapiAppointmentException : public MapiAppointment
{
public:
    MapiAppointmentException(MapiConnector2 *connection, const char *tallocName, MapiId &id, MapiAppointment &parent, AppointmentRecurrencePattern *pattern);

    /**
     * Fetch all properties.
     */
    virtual bool propertiesPull();

private:
    /**
     * Populate the object with property contents.
     */
    bool preparePayload();

    virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
    {
        Q_UNUSED(tags);
        Q_UNUSED(tagsAppended);
        Q_UNUSED(pullAll);

        return false;
    }

    MapiAppointment &m_parent;
    AppointmentRecurrencePattern *m_pattern;
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

bool ExCalResource::retrieveItem(const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts)
{
    Q_UNUSED(parts);

    // Eeeek. This is a bit racy, but hopefully good enough until we find out
    // what the rules really are.
    if (m_exceptionItems.size()) {
        deferTask();
        return true;
    }

    MapiAppointment *message = fetchItem<MapiAppointment>(itemOrig);
    if (!message) {
        return false;
    }

    // Create a clone of the passed in Item and fill it with the payload.
    message->setUid(itemOrig.remoteId());
    Akonadi::Item item(itemOrig);
    // TODO add further message properties.
    //item.setModificationTime(message->modified);
    item.setPayload<KCalCore::Event::Ptr>(KCalCore::Event::Ptr(message));

    // Notify Akonadi about the new data.
    itemRetrieved(item);

    // Start an asynchronous effort to create the exception items.
    foreach (MapiAppointmentException *exception, message->m_exceptions) {
        // An exception has the same UID as the original item.
        exception->setUid(item.remoteId());
        Akonadi::Item exceptionItem(m_itemMimeType);
        exceptionItem.setParentCollection(currentCollection());
        exceptionItem.setRemoteId(exception->id().toString());
        exceptionItem.setRemoteRevision(QString::number(1));
        exceptionItem.setPayload<KCalCore::Event::Ptr>(KCalCore::Event::Ptr(exception));
        m_exceptionItems.append(exceptionItem);
    }
    if (m_exceptionItems.size()) {
        QMetaObject::invokeMethod(this, "deleteExceptionItems", Qt::QueuedConnection);
    }
    return true;
}

/**
 * Delete the list of items we are about to create.
 *
 * Next state: @ref createExceptionItem().
 */
void ExCalResource::deleteExceptionItems()
{
    Akonadi::ItemDeleteJob *job = new Akonadi::ItemDeleteJob(m_exceptionItems);
    connect(job, SIGNAL(result(KJob *)), SLOT(createExceptionItem(KJob *)));
}

/**
 * Start the creation of a single exception item.
 *
 * Next state: @ref createExceptionItemDone().
 */
void ExCalResource::createExceptionItem(KJob *job)
{
    if (job->error()) {
        // Modify normal error reporting, since a delete can give us the
        // error "Unknown error. (No items found)".
        static QString noItems = QString::fromAscii("Unknown error. (No items found)");
        if (job->errorString() != noItems) {
            kError() << __FUNCTION__ << job->errorString();
        }
    }
    Akonadi::Item item = m_exceptionItems.first();
    m_exceptionItems.removeFirst();

    // Save the new item in Akonadi.
    kError() << __FUNCTION__ << "create" << item.remoteId() << "in" << item.parentCollection();
    Akonadi::ItemCreateJob *createJob = new Akonadi::ItemCreateJob(item, item.parentCollection());
    connect(createJob, SIGNAL(result(KJob *)), SLOT(createExceptionItemDone(KJob *)));
}

/**
 * Complete the creation of a single exception item.
 *
 * Next state: If there are more exception items, create the next item
 * @ref createExceptionItem(), otherwise we are done.
 */
void ExCalResource::createExceptionItemDone(KJob *job)
{
    if (job->error()) {
        kError() << __FUNCTION__ << job->errorString();
    }
    if (m_exceptionItems.size()) {
        // Go back and create the next item.
        createExceptionItem(job);
    }
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
        item.parentCollection().name() << "," << item.id() << "} = {" <<
        item.parentCollection().remoteId() << "," << item.remoteId() << "}";
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
#if 0
    KCalCore::Event::Ptr event = item.payload<KCalCore::Event::Ptr>();
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
    message->bodyText = event->description();
    //message->sender = event->organizer().name();
    message->location = event->location();
    if (event->alarms().count()) {
        KCalCore::Alarm::Ptr alarm = event->alarms().first();
        message->reminderSet = true;
        // TODO Maybe we should check which one is set and then use either the time or the delte
        // KDateTime reminder(message->reminderTime);
        // reminder.setTimeSpec( KDateTime::Spec(KDateTime::UTC) );
        // alarm->setTime( reminder );
        message->reminderDelta = alarm->startOffset() / -60;
    } else {
        message->reminderSet = false;
    }

    MapiRecipient att(MapiRecipient::Sender);
    att.name = event->organizer()->name();
    att.email = event->organizer()->email();
    message->addUniqueRecipient("event organiser", att);

    att.setType(MapiRecipient::To);
    foreach (KCalCore::Attendee::Ptr person, event->attendees()) {
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
#endif
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
    KCalCore::Event()
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

void MapiAppointment::ex2kcalRecurrency(AppointmentRecurrencePattern *pattern, KCalCore::Recurrence *kcal)
{
    kcal->clear();
    if (!pattern) {
        // No recurrency.
        return;
    }

    RecurrencePattern *ex = &pattern->RecurrencePattern;
    QString description;
    //debug() << "Calendar:" << ex->CalendarType;
    QBitArray days;
    if (ex->RecurFrequency == RecurFrequency_Daily && ex->PatternType == PatternType_Day) {
        kcal->setDaily(ex2kcalDaysFromMinutes(ex->Period));
        description = i18n("Every %1 days,", ex2kcalDaysFromMinutes(ex->Period));
    }
    else if (ex->RecurFrequency == RecurFrequency_Daily && ex->PatternType == PatternType_Week) {
        days = ex2kcalRecurrenceDays(M | Tu | W | Th | F);
        kcal->setWeekly(ex2kcalDaysFromMinutes(ex->Period) / 7, days, ex2kcalDayOfWeek(ex->FirstDOW));
        description = i18n("Every weekday,");
    }
    else if (ex->RecurFrequency == RecurFrequency_Weekly && ex->PatternType == PatternType_Week) {
        days = ex2kcalRecurrenceDays(ex->PatternTypeSpecific.WeekRecurrencePattern);
        kcal->setWeekly(ex->Period, days, ex2kcalDayOfWeek(ex->FirstDOW));
        description = i18n("Every %1 weeks on %2,", ex->Period, stringify(days));
    }
    else if (ex->RecurFrequency == RecurFrequency_Monthly) {
        kcal->setMonthly(ex->Period);
        switch (ex->PatternType)
        {
        case PatternType_Month:
        case PatternType_HjMonth:
            kcal->addMonthlyDate(ex->PatternTypeSpecific.Day);
            description = i18n("On the %1 day every %2 months,", ex->PatternTypeSpecific.Day,
                       ex->Period);
            break;
        case PatternType_MonthNth:
        case PatternType_HjMonthNth:
            days = ex2kcalRecurrenceDays(ex->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern);
            kcal->addMonthlyPos(ex->PatternTypeSpecific.MonthRecurrencePattern.N, days);
            description = i18n("On %1 of the %2 week every %3 months,", stringify(days),
                       ex->PatternTypeSpecific.MonthRecurrencePattern.N,
                       ex->Period);
            break;
        case PatternType_MonthEnd:
        case PatternType_HjMonthEnd:
            kcal->addMonthlyDate(ex->PatternTypeSpecific.Day);
            description = i18n("At the end (day %1) of every %2 month,", ex->PatternTypeSpecific.Day,
                       ex->Period);
            break;
        default:
            description = i18n("Unsupported monthly frequency with patterntype %1", ex->PatternType);
            break;
        }
    }
    else if (ex->RecurFrequency == RecurFrequency_Yearly) {
        kcal->setYearly(1);
        switch (ex->PatternType)
        {
        case PatternType_Month:
        case PatternType_HjMonth:
            kcal->addYearlyMonth(ex->Period);
            kcal->addYearlyDate(ex->PatternTypeSpecific.Day);
            description = i18n("Yearly, on the %1 day of the %2 month,", ex->PatternTypeSpecific.Day,
                       ex->Period);
            break;
        case PatternType_MonthNth:
        case PatternType_HjMonthNth:
            days = ex2kcalRecurrenceDays(ex->PatternTypeSpecific.MonthRecurrencePattern.WeekRecurrencePattern);
            kcal->addYearlyMonth(ex->Period);
            kcal->addYearlyPos(ex->PatternTypeSpecific.MonthRecurrencePattern.N, days);
            description = i18n("Yearly, the %1 day of the %2 month when it falls on %3,",
                       ex->PatternTypeSpecific.MonthRecurrencePattern.N,
                       ex->Period, stringify(days));
            break;
        default:
            description = i18n("Unsupported yearly frequency with patterntype %1", ex->PatternType);
            break;
        };
    } else {
        description = i18n("Unsupported frequency %1 with patterntype %2 combination", ex->RecurFrequency,
                   ex->PatternType);
    }

    kcal->setStartDateTime(ex2kcalTimes(ex->StartDate));
    switch (ex->EndType) {
    case END_AFTER_DATE:
        kcal->setEndDateTime(ex2kcalTimes(ex->EndDate));
        description += i18n(" from %1 to %2", stringify(kcal->startDateTime()),
                    stringify(kcal->endDateTime()));
        break;
    case END_AFTER_N_OCCURRENCES:
        kcal->setDuration(ex->OccurrenceCount);
        description += i18n(" from %1 for %2 occurrences", stringify(kcal->startDateTime()),
                    kcal->duration());
        break;
    case END_NEVER_END:
    case NEVER_END:
        description += i18n(" from %1, ending indefinitely", stringify(kcal->startDateTime()));
        break;
    default:
        description += i18n(" from %1, unsupported endtype %2", stringify(kcal->startDateTime()),
                    ex->EndType);
        break;
    }
    debug() << description;

    // We have dealt with the basic recurrence, now see what exceptions we have.
    for (int i = 0; i < pattern->ExceptionCount; i++) {
        MapiId exceptionId(m_id, (mapi_id_t)i);
        MapiAppointmentException *exception = new MapiAppointmentException(m_connection, "MapiAppointmentException", exceptionId, *this, pattern);

        if (!exception->propertiesPull()) {
            error() << "Error creating exception:" << exceptionId.toString();
            // Carry on regardless.
            delete exception;
            continue;
        }
        m_exceptions.append(exception);
    }
}

bool MapiAppointment::preparePayload()
{
    // Start with a clean slate.
    uint32_t sequence = 0;
    enum FreeBusyStatus busyStatus = olFree;
    QString location;
    QDateTime begin;
    QDateTime end;
    bool allDay = false;
    AppointmentStates state;
    ResponseStatus responseStatus = None;
    QString bodyText;
    QString bodyHtml;
    struct TimeZoneStruct *timezone = 0;
    AppointmentRecurrencePattern *pattern = 0;
    enum RecurFrequency recurrenceType = (enum RecurFrequency)0;
    QString messageClass;
    bool reminderSet = false;
    QDateTime reminderTime;
    uint32_t reminderDelta = 0;
    QString title;
    QDateTime modified;
    QDateTime created;

    // Walk through the properties and extract the values of interest. The
    // properties here should be aligned with the list pulled above.
    bool embeddedInBody = false;
    QString header;

    for (unsigned i = 0; i < m_propertyCount; i++) {
        MapiProperty property(m_properties[i]);

        switch (property.tag()) {
        case PidLidAppointmentSequence:
            sequence = property.value().toUInt();
            break;
        case PidLidBusyStatus:
            busyStatus = (enum FreeBusyStatus)property.value().toUInt();
            break;
        case PidLidLocation:
            location = property.value().toString();
            break;
        case PidLidAppointmentStartWhole:
            begin = property.value().toDateTime();
            break;
        case PidLidAppointmentEndWhole:
            end = property.value().toDateTime();
            break;
        case PidLidAppointmentSubType:
            allDay = property.value().toUInt() != 0;
            break;
        case PidLidAppointmentStateFlags:
            state = AppointmentStates(property.value().toUInt());
            break;
        case PidLidResponseStatus:
            responseStatus = (ResponseStatus)property.value().toUInt();
            break;
        case PidTagBody:
            bodyText = property.value().toString();
            break;
        case PidTagHtml:
            bodyHtml = property.value().toString();
            break;
        case PidLidTimeZoneStruct:
            timezone = get_TimeZoneStruct(ctx(), &m_properties[i].value.bin);
            break;
        case PidLidAppointmentRecur:
            pattern = get_AppointmentRecurrencePattern(ctx(), &m_properties[i].value.bin);
            break;
        case PidLidRecurrenceType:
            recurrenceType = (enum RecurFrequency)property.value().toInt();
            break;
        case PidTagMessageClass:
            // Sanity check the message class.
            messageClass = property.value().toString();
            if (!messageClass.startsWith(QLatin1String("IPM.Appointment"))) {
                if (!messageClass.startsWith(QLatin1String("IPM.Note"))) {
                    error() << "retrieved item is not an appointment:" << messageClass;
                    return false;
                } else {
                    embeddedInBody = true;
                }
            }
            break;
        case PidLidReminderSet:
            reminderSet = property.value().toInt();
            break;
        case PidLidReminderSignalTime:
            reminderTime = property.value().toDateTime();
            break;
        case PidLidReminderDelta:
            reminderDelta = property.value().toInt();
            break;
        case PidTagConversationTopic:
            title = property.value().toString();
            break;
        case PidTagLastModificationTime:
            modified = property.value().toDateTime();
            break;
        case PidTagCreationTime:
            created = property.value().toDateTime();
            break;
        case PidTagTransportMessageHeaders:
            header = property.value().toString();
            break;
        default:
            // Handle oversize objects.
            if (MAPI_E_NOT_ENOUGH_MEMORY == property.value().toInt()) {
                switch (property.tag()) {
                case PidTagBody_Error:
                    if (!streamRead(&m_object, PidTagBody, CODEPAGE_UTF16, bodyText)) {
                        return false;
                    }
                    break;
                case PidTagHtml_Error:
                    if (!streamRead(&m_object, PidTagHtml, CODEPAGE_UTF16, bodyHtml)) {
                        return false;
                    }
                    break;
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
        //  Microsoft Mail Internet Headers Version 2.0
        //  BEGIN:VCALENDAR
        //  PRODID:-//K Desktop Environment//NONSGML libkcal 4.3//EN
        //  VERSION:2.0
        //  BEGIN:VEVENT
        //  DTSTAMP:20100420T092856Z
        //  ORGANIZER;CN="xxx yyy (zzz)":MAILTO:zzz@foo.com
        //  X-UID: 0
        //
        // and the rest in the body:
        //
        //  ATTENDEE;CN="aaa bbb (ccc)";RSVP=TRUE;PARTSTAT=NEEDS-ACTION;
        //  ...
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
        bodyText = header + bodyText;
        if (bodyText.isEmpty()) {
            error() << "retrieved content is not an appointment";
            return false;
        }
        return true;
    }

    // Now set all the properties onto the item.
    switch (busyStatus) {
    case olFree:
        setTransparency(Event::Transparent);
        break;
    default:
        setTransparency(Event::Opaque);
        break;
    }
    setLocation(location);
    setDtStart(KDateTime(begin));
    setDtEnd(KDateTime(end));
    setAllDay(allDay);
    if (state.testFlag(Canceled)) {
        // Just mark this entry as cancelled.
        setStatus(StatusCanceled);
    } else {
        // If this came from somebody else, the user may have to do something.
        if (state.testFlag(Received)) {
            switch (responseStatus) {
            case None:
                setStatus(StatusNone);
                break;
            case Organized:
                // TODO I'm not sure what this means exactly?
                setStatus(StatusNeedsAction);
                break;
            case Tentative:
                setStatus(StatusTentative);
                break;
            case Accepted:
                setStatus(StatusConfirmed);
                break;
            case Declined:
                setStatus(StatusCanceled);
                break;
            case NotResponded:
                setStatus(StatusNeedsAction);
                break;
            }
        } else {
            // This did not come from anybody else.
            setStatus(StatusConfirmed);
        }
    }
    setDescription(bodyText);
    setAltDescription(bodyHtml);
    // TODO timezone
    if (recurrenceType != 0) {
        if (!pattern) {
            // This should not happen. PidLidRecurrenceType says this is a
            // recurring event, so why is there no PidLidAppointmentRecur???
            error() << "missing pattern for recurrenceType:" << recurrenceType;
            recurrenceType = (enum RecurFrequency)0;
        } else {
            error() << " got recurrence**********";
            ex2kcalRecurrency(pattern, recurrence());
        }
    }
    if (reminderSet) {
        KCalCore::Alarm::Ptr alarm(new KCalCore::Alarm(dynamic_cast<KCalCore::Incidence*>(this)));
        // TODO Maybe we should check which one is set and then use either the time or the delte
        // KDateTime reminder(reminderTime);
        // reminder.setTimeSpec(KDateTime::Spec(KDateTime::UTC));
        // alarm->setTime(reminder);
        alarm->setStartOffset(KCalCore::Duration(reminderDelta * -60));
        alarm->setEnabled(true);
        addAlarm(alarm);
    }
    setSummary(title);
    setLastModified(KDateTime(modified));
    setCreated(KDateTime(created));
    foreach (MapiRecipient recipient, recipients()) {
        if (recipient.type() == MapiRecipient::ReplyTo) {
            KCalCore::Person::Ptr person(new KCalCore::Person(recipient.name, recipient.email));
            setOrganizer(person);
        } else {
            KCalCore::Attendee::Ptr person(new KCalCore::Attendee(recipient.name, recipient.email));
            addAttendee(person);
        }
    }
    return true;
}

bool MapiAppointment::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
    /**
     * The list of tags used to fetch an Appointment, based on [MS-OXOCAL].
     */
    static unsigned ourTagList[] = {
        // 2.2.1.1
        PidLidAppointmentSequence,
        // 2.2.1.2
        PidLidBusyStatus,
        // 2.2.1.3
        //PidLidAppointmentAuxiliaryFlags,
        // 2.2.1.4
        PidLidLocation,
        // 2.2.1.5
        PidLidAppointmentStartWhole,
        // 2.2.1.6
        PidLidAppointmentEndWhole,
        // 2.2.1.7
        //PidLidAppointmentDuration,
        // 2.2.1.8
        //PidNameKeywords,
        // 2.2.1.9
        PidLidAppointmentSubType,
        // 2.2.1.10
        PidLidAppointmentStateFlags,
        // 2.2.1.11
        PidLidResponseStatus,
        // 2.2.1.12
        //PidLidRecurring,
        // 2.2.1.13
        //PidLidIsRecurring,
        // 2.2.1.14
        //PidLidClipStart,
        // 2.2.1.15
        //PidLidClipEnd,
        // 2.2.1.16
        //PidLidAllAttendeesString,
        // 2.2.1.17
        //PidLidToAttendeesString,
        // 2.2.1.18
        //PidLidCcAttendeesString,
        // 2.2.1.19
        //PidLidNonSendableTo,
        // 2.2.1.20
        //PidLidNonSendableCc,
        // 2.2.1.21
        //PidLidNonSendableBcc,
        // 2.2.1.22
        //PidLidNonSendToTrackStatus,
        // 2.2.1.23
        //PidLidNonSendCcTrackStatus,
        // 2.2.1.24
        //PidLidNonSendBccTrackStatus,
        // 2.2.1.25
        //PidLidAppointmentUnsendableRecipients,
        // 2.2.1.26
        //PidLidAppointmentNotAllowPropose,
        // 2.2.1.27
        //PidLidGlobalObjectId,
        // 2.2.1.28
        //PidLidCleanGlobalObjectId,
        // 2.2.1.29
        //PidTagOwnerAppointmentId,
        // 2.2.1.30
        //PidTagStartDate,
        // 2.2.1.31
        //PidTagEndDate,
        // 2.2.1.32
        //PidLidCommonStart,
        // 2.2.1.33
        //PidLidCommonEnd,
        // 2.2.1.34
        //PidLidOwnerCriticalChange,
        // 2.2.1.35
        //PidLidIsException,
        // 2.2.1.36
        //PidTagResponseRequested,
        // 2.2.1.37
        //PidTagReplyRequested,
        // 2.2.1.38 Best Body Properties
        PidTagBody,
        PidTagHtml,
        // 2.2.1.39
        PidLidTimeZoneStruct,
        // 2.2.1.40
        //PidLidTimeZoneDescription,
        // 2.2.1.41
        //PidLidAppointmentTimeZoneDefinitionRecur,
        // 2.2.1.42
        //PidLidAppointmentTimeZoneDefinitionStartDisplay,
        // 2.2.1.43
        //PidLidAppointmentTimeZoneDefinitionEndDisplay,
        // 2.2.1.44
        PidLidAppointmentRecur,
        // 2.2.1.45
        PidLidRecurrenceType,
        // 2.2.1.46
        //PidLidRecurrencePattern,
        // 2.2.1.47
        //PidLidLinkedTaskItems,
        // 2.2.1.48
        //PidLidMeetingWorkspaceUrl,
        // 2.2.1.49
        //PidTagIconIndex Property
        // 2.2.2.1
        PidTagMessageClass,
        // 2.2.3 Appointment-specific, nothing needed.
        // TODO 2.2.4 through 2.2.9 Meeting-specific.
        // TODO 2.2.10 Exception objects.
        // 2.2.11 Calendar folder, nothing needed.
        // TODO 2.2.12 Delegates.
        // [MS-OXORMDR] section 2.2.1.1
        PidLidReminderSet,
        // [MS-OXORMDR] section 2.2.1.2
        PidLidReminderSignalTime,
        // [MS-OXORMDR] section 2.2.1.3
        PidLidReminderDelta,
        // Other
        PidTagConversationTopic,
        PidTagLastModificationTime,
        PidTagCreationTime,
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
    if (!preparePayload()) {
        return false;
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
#if 0
    if (!propertyWrite(PidTagConversationTopic, title)) {
        return false;
    }
    if (!propertyWrite(PidTagBody, bodyText)) {
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
    if (!propertyWrite(PidLidReminderSet, reminderSet)) {
        return false;
    }
    if (reminderSet) {
        if (!propertyWrite(PidLidReminderSignalTime, reminderTime)) {
            return false;
        }
        if (!propertyWrite(PidLidReminderDelta, reminderDelta)) {
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
            debug() << "missing pattern in message"<<messageID<<"in folder"<<folderID;
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
#endif
    return true;
}

MapiAppointmentException::MapiAppointmentException(MapiConnector2 *connection,
                                                   const char *tallocName,
                                                   MapiId &id, MapiAppointment &parent,
                                                   AppointmentRecurrencePattern *pattern) :
    MapiAppointment(connection, tallocName, id),
    m_parent(parent),
    m_pattern(pattern)
{
}

bool MapiAppointmentException::preparePayload()
{
    ExceptionInfo *e = &m_pattern->ExceptionInfo[m_id.second];
    ExtendedException *ee = &m_pattern->ExtendedException[m_id.second];

    // Set the main properties from the parent or the always-present parts of
    // the exception information.
    enum FreeBusyStatus busyStatus = (m_parent.transparency() == Event::Transparent) ? olFree : olBusy;
    QString location = m_parent.location();
    KDateTime begin = ex2kcalTimes(e->StartDateTime);
    KDateTime end = ex2kcalTimes(e->EndDateTime);
    bool allDay = m_parent.allDay();
    AppointmentStates state;
    bool reminderSet = m_parent.alarms().size() > 0;
    uint32_t reminderDelta = reminderSet ? (m_parent.alarms().first()->startOffset().asSeconds() / -60) : 0;
    QString title = m_parent.summary();
    uint32_t changeHighlight = 0;
    bool attachment = false;
    KDateTime originalBegin = ex2kcalTimes(e->OriginalStartDate);
    OverrideFlags overrideFlags = e->OverrideFlags;

    QString description;
    description += i18n("\nException%1 for %2 to be from %3 to %4", m_id.second, stringify(originalBegin), stringify(begin), stringify(end));
#if 0
    // ExtendedException has Unicode strings, but needs Openchange
    // bug #391 to be fixed.
    if (ee->WriterVersion2 >= 0x00003009) {
        changeHighlight = ee->ChangeHighlight.ChangeHighlight.ChangeHighlightValue;
        description += i18n("\n    ChangeHighlight %1", changeHighlight);
    }
    if (overrideFlags & ARO_SUBJECT) {
        title.setUtf16(ee->Subject.Msg.Msg, ee->Subject.Msg.Length);
        description += i18n("\n    Subject %1", title);
    }
    if (overrideFlags & ARO_LOCATION) {
        location.setUtf16(ee->Location.Msg.Msg, ee->Location.Msg.Length);
        description += i18n("\n    Location %1", location);
    }
#else
    // Non-unicode support only, from the ExceptionInfo.
    if (overrideFlags & ARO_SUBJECT) {
        title = QString::fromLatin1((const char *)e->Subject.Msg.msg, e->Subject.Msg.msgLength2);
        description += i18n("\n    Subject %1", title);
    }
    if (overrideFlags & ARO_LOCATION) {
        location = QString::fromLatin1((const char *)e->Location.Msg.msg, e->Location.Msg.msgLength2);
        description += i18n("\n    Location %1", location);
    }
#endif
    if (overrideFlags & ARO_MEETINGTYPE) {
        state = AppointmentStates(e->MeetingType.Value);
    }
    if (overrideFlags & ARO_REMINDERDELTA) {
        reminderDelta = e->ReminderDelta.Value;
    }
    if (overrideFlags & ARO_REMINDER) {
        reminderSet = e->ReminderSet.Value != 0;
    }
    if (overrideFlags & ARO_BUSYSTATUS) {
        busyStatus = (enum FreeBusyStatus)e->BusyStatus.Value;
    }
    if (overrideFlags & ARO_ATTACHMENT) {
        attachment = e->Attachment.Value != 0;
    }
    if (overrideFlags & ARO_SUBTYPE) {
        allDay = e->SubType.Value != 0;
    }
    description += i18n("\n    ChangeHighlight %1, state %2, reminderDelta %3, reminderSet %4, busyStatus %5, attachment %6, allDay %7",
                changeHighlight, state, reminderDelta, reminderSet, busyStatus, attachment, allDay);
    debug() << description;

    // Now set all the properties onto the item. Any items not specified by the
    // exception are just copied from the parent.
    switch (busyStatus) {
    case olFree:
        setTransparency(Event::Transparent);
        break;
    default:
        setTransparency(Event::Opaque);
        break;
    }
    setLocation(location);
    setDtStart(begin);
    setDtEnd(end);
    setAllDay(allDay);
    setStatus(m_parent.status());
    setDescription(m_parent.description());
    setAltDescription(m_parent.altDescription());
    if (reminderSet) {
        KCalCore::Alarm::Ptr alarm(new KCalCore::Alarm(dynamic_cast<KCalCore::Incidence*>(this)));
        // TODO Maybe we should check which one is set and then use either the time or the delte
        // KDateTime reminder(reminderTime);
        // reminder.setTimeSpec(KDateTime::Spec(KDateTime::UTC));
        // alarm->setTime(reminder);
        alarm->setStartOffset(KCalCore::Duration(reminderDelta * -60));
        alarm->setEnabled(true);
        addAlarm(alarm);
    } else {
        // The parent will have at most one alarm.
        if (m_parent.alarms().size()) {
            addAlarm(m_parent.alarms().first());
        }
    }
    setSummary(title);
    setLastModified(m_parent.lastModified());
    setCreated(m_parent.created());
    setOrganizer(m_parent.organizer());
    foreach (const KCalCore::Attendee::Ptr person, m_parent.attendees()) {
        addAttendee(person);
    }

    // Create relationship with the parent recurrence: the exception has the same
    // UID as the original event, and a recurrenceId of the original item instance
    // to be overridden.
    setRecurrenceId(originalBegin);
    KCalCore::RecurrenceRule *r = new KCalCore::RecurrenceRule();
    r->setAllDay(allDay);
    r->setStartDt(begin);
    r->setEndDt(end);
    m_parent.recurrence()->addExDate(originalBegin.date());
    m_parent.recurrence()->addExDateTime(originalBegin);
    m_parent.recurrence()->addExRule(r);
    return true;
}

bool MapiAppointmentException::propertiesPull()
{
    return preparePayload();
}

AKONADI_RESOURCE_MAIN( ExCalResource )

#include "excalresource.moc"
