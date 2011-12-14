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

#include "mapiconnector2.h"

#include <QDebug>
#include <QStringList>
#include <QDir>
#include <QMessageBox>
#include <QRegExp>
#include <QVariant>

/**
 * Display ids in the same format we use when stored in Akonadi.
 */
#define ID_FORMAT 0, 36

#define CASE_PREFER_UNICODE(unicode, lvalue, rvalue) \
case unicode ## _string8: \
	if (lvalue.isEmpty()) { \
		lvalue = rvalue; \
	} \
	break; \
case unicode: \
	lvalue = rvalue; \
	break;

/**
 * Set this to 1 to pull all the properties for an appointment or a note 
 * respectively, e.g. to see what a server has available.
 */
#define DEBUG_APPOINTMENT_PROPERTIES 0
#define DEBUG_NOTE_PROPERTIES 1

/**
 * Map all MAPI errors to strings. Note that all MAPI error handling 
 * assumes that MAPI_E_SUCCESS == 0!
 */
static QString mapiError()
{
	int code = GetLastError();
#define STR(e_code) \
if (e_code == code) { return QString::fromLatin1(#e_code); }
	STR(MAPI_E_SUCCESS);
	STR(MAPI_E_INTERFACE_NO_SUPPORT);
	STR(MAPI_E_CALL_FAILED);
	STR(MAPI_E_NO_SUPPORT);
	STR(MAPI_E_BAD_CHARWIDTH);
	STR(MAPI_E_STRING_TOO_LONG);
	STR(MAPI_E_UNKNOWN_FLAGS);
	STR(MAPI_E_INVALID_ENTRYID);
	STR(MAPI_E_INVALID_OBJECT);
	STR(MAPI_E_OBJECT_CHANGED);
	STR(MAPI_E_OBJECT_DELETED);
	STR(MAPI_E_BUSY);
	STR(MAPI_E_NOT_ENOUGH_DISK);
	STR(MAPI_E_NOT_ENOUGH_RESOURCES);
	STR(MAPI_E_NOT_FOUND);
	STR(MAPI_E_VERSION);
	STR(MAPI_E_LOGON_FAILED);
	STR(MAPI_E_SESSION_LIMIT);
	STR(MAPI_E_USER_CANCEL);
	STR(MAPI_E_UNABLE_TO_ABORT);
	STR(MAPI_E_NETWORK_ERROR);
	STR(MAPI_E_DISK_ERROR);
	STR(MAPI_E_TOO_COMPLEX);
	STR(MAPI_E_BAD_COLUMN);
	STR(MAPI_E_EXTENDED_ERROR);
	STR(MAPI_E_COMPUTED);
	STR(MAPI_E_CORRUPT_DATA);
	STR(MAPI_E_UNCONFIGURED);
	STR(MAPI_E_FAILONEPROVIDER);
	STR(MAPI_E_UNKNOWN_CPID);
	STR(MAPI_E_UNKNOWN_LCID);
	STR(MAPI_E_PASSWORD_CHANGE_REQUIRED);
	STR(MAPI_E_PASSWORD_EXPIRED);
	STR(MAPI_E_INVALID_WORKSTATION_ACCOUNT);
	STR(MAPI_E_INVALID_ACCESS_TIME);
	STR(MAPI_E_ACCOUNT_DISABLED);
	STR(MAPI_E_END_OF_SESSION);
	STR(MAPI_E_UNKNOWN_ENTRYID);
	STR(MAPI_E_MISSING_REQUIRED_COLUMN);
	STR(MAPI_E_BAD_VALUE);
	STR(MAPI_E_INVALID_TYPE);
	STR(MAPI_E_TYPE_NO_SUPPORT);
	STR(MAPI_E_UNEXPECTED_TYPE);
	STR(MAPI_E_TOO_BIG);
	STR(MAPI_E_DECLINE_COPY);
	STR(MAPI_E_UNEXPECTED_ID);
	STR(MAPI_E_UNABLE_TO_COMPLETE);
	STR(MAPI_E_TIMEOUT);
	STR(MAPI_E_TABLE_EMPTY);
	STR(MAPI_E_TABLE_TOO_BIG);
	STR(MAPI_E_INVALID_BOOKMARK);
	STR(MAPI_E_WAIT);
	STR(MAPI_E_CANCEL);
	STR(MAPI_E_NOT_ME);
	STR(MAPI_E_CORRUPT_STORE);
	STR(MAPI_E_NOT_IN_QUEUE);
	STR(MAPI_E_NO_SUPPRESS);
	STR(MAPI_E_COLLISION);
	STR(MAPI_E_NOT_INITIALIZED);
	STR(MAPI_E_NON_STANDARD);
	STR(MAPI_E_NO_RECIPIENTS);
	STR(MAPI_E_SUBMITTED);
	STR(MAPI_E_HAS_FOLDERS);
	STR(MAPI_E_HAS_MESAGES);
	STR(MAPI_E_FOLDER_CYCLE);
	STR(MAPI_E_LOCKID_LIMIT);
	STR(MAPI_E_AMBIGUOUS_RECIP);
	STR(MAPI_E_NAMED_PROP_QUOTA_EXCEEDED);
	STR(MAPI_E_NOT_IMPLEMENTED);
	STR(MAPI_E_NO_ACCESS);
	STR(MAPI_E_NOT_ENOUGH_MEMORY);
	STR(MAPI_E_INVALID_PARAMETER);
	STR(MAPI_E_RESERVED);
	return QString::fromLatin1("MAPI_E_0x%1").arg((unsigned)code, 0, 16);
}

static QDateTime convertSysTime(const FILETIME& filetime)
{
  NTTIME nt_time = filetime.dwHighDateTime;
  nt_time = nt_time << 32;
  nt_time |= filetime.dwLowDateTime;
  QDateTime kdeTime;
  time_t unixTime = nt_time_to_unix(nt_time);
  kdeTime.setTime_t(unixTime);
  //kDebug() << "unix:"<<unixTime << "time:"<<kdeTime.toString() << "local:"<<kdeTime.toLocalZone();
  return kdeTime;
}

static int profileSelectCallback(struct SRowSet *rowset, const void* /*private_var*/)
{
	qCritical() << "Found more than 1 matching users -> cancel";

	//  TODO Some sort of handling would be needed here
	return rowset->cRows;
}

/**
 * A very simple wrapper around a property.
 */
class MapiProperty : private SPropValue
{
public:
	MapiProperty(SPropValue &property);

	/**
	 * Get the value of the property in a nice typesafe wrapper.
	 */
	QVariant value() const;

	/**
	 * Get the string equivalent of a property, e.g. for display purposes.
	 * We take care to hex-ify GUIDs and other byte arrays, and lists of
	 * the same.
	 */
	QString toString() const;

	/**
	 * Return the integer tag.
	 * 
	 * To convert this into a name, @ref MapiObject::tagName().
	 */
	int tag() const;

private:
	SPropValue &m_property;
};

MapiAppointment::MapiAppointment(MapiConnector2 *connector, const char *tallocName, mapi_id_t folderId, mapi_id_t id) :
	MapiMessage(connector, tallocName, folderId, id)
{
}

QDebug MapiAppointment::debug() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT)) /*<< title*/;
}

QDebug MapiAppointment::error() const
{
	static QString prefix = QString::fromAscii("MapiAppointment: %1/%2:");
	return MapiObject::error(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT)) /*<< title*/;
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

/**
 * Add an attendee to the list. We:
 *
 *  1. Check whether the name has an embedded email address, and if it does
 *     use it (if it is better than the one we have).
 *
 *  2. If the name matches an existing entry in the list, just pick the
 *     better email address.
 *
 *  3. If the email matches, just pick the better name.
 *
 * Otherwise, we have a new entry.
 */
void MapiAppointment::addUniqueAttendee(Attendee candidate)
{
	// See if we can deduce a better email address from the name than we 
	// have in the explicit value. Look for the last possible starting 
	// delimiter, and work forward from there. Thus:
	//
	//       "blah (blah) <blah> <result>"
	//
	// should return "result". Note that we don't remove this from the name
	// so as to give the resolution process as much to work with as possible.
	static QRegExp firstRE(QString::fromAscii("[(<]"));
	static QRegExp lastRE(QString::fromAscii("[)>]"));

	int first = candidate.name.lastIndexOf(firstRE);
	int last = candidate.name.indexOf(lastRE, first);

	if ((first > -1) && (last > first + 1)) {
		QString mid = candidate.name.mid(first + 1, last - first - 1);
		if (isGoodEmailAddress(candidate.email) < isGoodEmailAddress(mid)) {
			candidate.email = mid;
		}
	}

	for (int i = 0; i < attendees.size(); i++) {
		Attendee &entry = attendees[i];

		// If we find a name match, fill in a missing email if we can.
		if (!candidate.name.isEmpty() && (entry.name == candidate.name)) {
			if (isGoodEmailAddress(entry.email) < isGoodEmailAddress(candidate.email)) {
				entry.email = candidate.email;
				return;
			}
		}

		// If we find an email match, fill in a missing name if we can.
		if (!candidate.email.isEmpty() && (entry.email == candidate.email)) {
			if (entry.name.length() < candidate.name.length()) {
				entry.name = candidate.name;
				return;
			}
		}
	}

	// Add the entry if it did not match.
	attendees.append(candidate);
}

unsigned MapiAppointment::isGoodEmailAddress(QString &email)
{
	static QChar at(QChar::fromAscii('@'));
	static QChar x500Prefix(QChar::fromAscii('/'));

	// Anything is better than an empty address!
	if (email.isEmpty()) {
		return 0;
	}

	// KDEPIM does not currently handle anything like this:
	//
	// "/O=CISCO SYSTEMS/OU=FIRST ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=name"
	//
	// but for the user, it is still better than nothing, so we give it a
	// low priority.
	if (email[0] == x500Prefix) {
		return 1;
	}

	// An @ sign is better than no @ sign.
	if (!email.contains(at)) {
		return 2;
	} else {
		return 3;
	}
}

/**
 * We collect recipients as well as properties. The recipients are pulled from
 * multiple sources, but need to go through a resolution process to fix them up.
 * The duplicates will disappear as part of the resolution process. The logic
 * looks like this:
 *
 *  1. Read the RecipientsTable.
 *
 *  2. Add in the contents of the DisplayTo tag (Exchange can return records
 *     from Step 1 with no name or email).
 *
 *  3. Add in the sent-representing values (also not in Step 1, and not clear
 *     if it is guaranteed to be in step 2 as well).
 *
 * That results in a whole lot of duplication as well as bringing in the missing
 * items. So then, there is a whole lot of work to resolve things:
 *
 *  - For anything we get from 1, 2 and 3 (mostly 2 and 3, but also 1) try hard
 *    to get an email address. Whenever I find one, I compare the "quality" of
 *    the address against what we already have, and keep the best one.
 *
 *  - For all but the best quality addresses, call ResolveNames. Again, keep
 *    the best value seen.
 *
 *  - For those that are left, wing-it.
 *
 * There is also a load of cruft data elimination along the way.
 */
bool MapiAppointment::propertiesPull()
{
// TEST -START-  Try an easier approach to find all the attendees
	struct ReadRecipientRow * recipientTable = 0x0;
	uint8_t recCount;
	ReadRecipients(d(), 0, &recCount, &recipientTable);
	qDebug()<< "number of recipients:"<<recCount;
	for (int x=0; x<recCount; x++) {
		struct ReadRecipientRow * recipient = &recipientTable[x];

		QString recipientName;
		QString recipientEmail;
		
		// Found this "intel" in MS-OXCDATA 2.8.3.1
		bool hasDisplayName = (recipient->RecipientRow.RecipientFlags&0x0010)==0x0010;
		if (hasDisplayName) {
			recipientName = QString::fromUtf8(recipient->RecipientRow.DisplayName.lpszW);
		}
		bool hasEmailAddress = (recipient->RecipientRow.RecipientFlags&0x0008)==0x0008;
		if (hasEmailAddress) {
			recipientEmail = QString::fromUtf8(recipient->RecipientRow.EmailAddress.lpszW);
		}

		uint16_t type = recipient->RecipientRow.RecipientFlags&0x0007;
		switch (type) {
			case 0x1: // => X500DN
				// we need to resolve this recipient's data
				// TODO for now we just copy the user's account name to the recipientEmail
				recipientEmail = QString::fromLocal8Bit(recipient->RecipientRow.X500DN.recipient_x500name);
				break;
			default:
				// don't see any need to evaluate anything but the "X500DN" case
				break;
		}

		qDebug()<< "recipient["<<x<<"] type:"<< recipient->RecipientType
				<< "flags:" << recipient->RecipientRow.RecipientFlags
				<< "name:" << recipientName
				<< "email:" << recipientEmail;
	}
// TEST -END-

	// Start with a clean slate.
	reminderActive = false;
	attendees.clear();

	// Step 1. Add all the recipients from the actual table.
	if (!MapiMessage::recipientsPull()) {
		return false;
	}
	for (unsigned i = 0; i < recipientCount(); i++) {
		addUniqueAttendee(Attendee(recipientAt(i)));
	}
#if (DEBUG_APPOINTMENT_PROPERTIES)
	if (!MapiMessage::propertiesPull()) {
		return false;
	}
#else
	QVector<int> readTags;
	readTags.append(PidTagMessageClass);
	readTags.append(PidTagDisplayTo);
	readTags.append(PidTagConversationTopic);
	readTags.append(PidTagBody);
	readTags.append(PidTagLastModificationTime);
	readTags.append(PidTagCreationTime);
	readTags.append(PidTagStartDate);
	readTags.append(PidTagEndDate);
	readTags.append(PidLidLocation);
	readTags.append(PidLidReminderSet);
	readTags.append(PidLidReminderSignalTime);
	readTags.append(PidLidReminderDelta);
	readTags.append(PidLidRecurrenceType);
	readTags.append(PidLidAppointmentRecur);

	readTags.append(PidTagSentRepresentingEmailAddress);
	readTags.append(PidTagSentRepresentingEmailAddress_string8);

	readTags.append(PidTagSentRepresentingName);
	readTags.append(PidTagSentRepresentingName_string8);

	readTags.append(PidTagSentRepresentingSimpleDisplayName);
	readTags.append(PidTagSentRepresentingSimpleDisplayName_string8);

	readTags.append(PidTagOriginalSentRepresentingEmailAddress);
	readTags.append(PidTagOriginalSentRepresentingEmailAddress_string8);

	readTags.append(PidTagOriginalSentRepresentingName);
	readTags.append(PidTagOriginalSentRepresentingName_string8);
	if (!MapiMessage::propertiesPull(readTags)) {
		return false;
	}
#endif

	QStringList displayTo;
	Attendee displayed;
	unsigned recurrenceType = 0;
	RecurrencePattern *pattern = 0;
	Attendee sentRepresenting;
	Attendee originalSentRepresenting;

	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		switch (property.tag()) {
			// Sanity check the message class.
			if (QLatin1String("IPM.Appointment") != property.value().toString()) {
				// this one is not an appointment
				return false;
			}
			case PidTagDisplayTo:
			// Step 2. Add the DisplayTo items.
			//
			// Astonishingly, when we call recipientsPull() later,
			// the results can contain entries with missing name (!)
			// and email values. Reading this property is a pathetic
			// workaround.
			displayTo = property.value().toString().split(QChar::fromAscii(';'));
			foreach (QString name, displayTo) {
				displayed.name = name.trimmed();
				addUniqueAttendee(displayed);
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
		CASE_PREFER_UNICODE(PidTagSentRepresentingEmailAddress, sentRepresenting.email, property.value().toString())
		CASE_PREFER_UNICODE(PidTagSentRepresentingName, sentRepresenting.name, property.value().toString())
		CASE_PREFER_UNICODE(PidTagSentRepresentingSimpleDisplayName, sentRepresenting.name, property.value().toString())
		CASE_PREFER_UNICODE(PidTagOriginalSentRepresentingEmailAddress, originalSentRepresenting.email, property.value().toString())
		CASE_PREFER_UNICODE(PidTagOriginalSentRepresentingName, originalSentRepresenting.name, property.value().toString())
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

	// Step 3. Add the sent-on-behalf-of (not the sender!).
	if (!originalSentRepresenting.name.isEmpty()) {
		sender = originalSentRepresenting.name;
		addUniqueAttendee(originalSentRepresenting);
		attendees.append(originalSentRepresenting);
	}
	if (!sentRepresenting.name.isEmpty()) {
		sender = sentRepresenting.name;
		addUniqueAttendee(sentRepresenting);
	}

	// We have all the attendees; find any that need resolution.
	QList<int> needingResolution;
	for (int i = 0; i < attendees.size(); i++) {
		Attendee &attendee = attendees[i];
		static QString perfectForm = QString::fromAscii("foo@foo");
		static unsigned perfect = isGoodEmailAddress(perfectForm);

		// If we find a missing/incomplete email, it needs resolution.
		if (isGoodEmailAddress(attendee.email) < perfect) {
			needingResolution << i;
		}
	}
	debug() << "attendees needing primary resolution:" << needingResolution.size() << "from a total:" << attendees.size();

	// Primary resolution is to ask Exchange to resolve the names.
	struct PropertyTagArray_r *statuses = NULL;
	struct SRowSet *results = NULL;
	if (needingResolution.size() > 0) {
		// Fill an array with the names we need to resolve. We will do a Unicode
		// lookup, so use UTF8.
		const char *names[needingResolution.size() + 1];
		unsigned j = 0;
		foreach (int i, needingResolution) {
			Attendee &attendee = attendees[i];
			QByteArray utf8(attendee.name.toUtf8());

			utf8.append('\0');
			names[j] = talloc_strdup(ctx(), utf8.data());
			j++;
		}
		names[j] = 0;
		struct SPropTagArray *tags = NULL;
		tags = set_SPropTagArray(ctx(), 10, 
					PidTag7BitDisplayName,         PidTagDisplayName,         PidTagRecipientDisplayName, 
					PidTag7BitDisplayName_string8, PidTagDisplayName_string8, PidTagRecipientDisplayName_string8,
					PidTagPrimarySmtpAddress,
					PidTagPrimarySmtpAddress_string8,
					0x6001001f,
					0x60010018);

		// Server round trip here!
		if (!m_connection->resolveNames(names, tags, &results, &statuses)) {
			return false;
		}
	}
	if (results) {
		// Walk the returned results. Every request has a status, but
		// only resolved items also have a row of results.
		//
		// As we do the walk, we trim the needingResolution array so
		// that when we are done with this loop, it only contains
		// entries which need more work.
		for (unsigned i = 0, unresolveds = 0; i < statuses->cValues; i++) {
			Attendee &attendee = attendees[needingResolution.at(unresolveds)];

			if (MAPI_RESOLVED == statuses->aulPropTag[i]) {
				struct SRow &recipient = results->aRow[i - unresolveds];
				QString name, email;
				for (unsigned j = 0; j < recipient.cValues; j++) {
					MapiProperty property(recipient.lpProps[j]);

					// Note that the set of properties fetched here must be aligned
					// with those fetched in MapiMessage::recipientAt(), as well
					// as the array above.
					switch (property.tag()) {
					CASE_PREFER_UNICODE(PidTag7BitDisplayName, name, property.value().toString())
					CASE_PREFER_UNICODE(PidTagDisplayName, name, property.value().toString())
					CASE_PREFER_UNICODE(PidTagRecipientDisplayName, name, property.value().toString())
					CASE_PREFER_UNICODE(PidTagPrimarySmtpAddress, email, property.value().toString())
					case 0x6001001e:
						if (email.isEmpty()) {
							email = property.value().toString();
						}
						break;
					case 0x6001001f:
						email = property.value().toString();
						break;
					default:
						break;
					}
				}

				// A resolved value is better than an unresolved one.
				if (!name.isEmpty()) {
					attendee.name = name;
				}
				if (isGoodEmailAddress(attendee.email) < isGoodEmailAddress(email)) {
					attendee.email = email;
				}
				needingResolution.removeAt(unresolveds);
			} else {
				unresolveds++;
			}
		}
	}
	MAPIFreeBuffer(results);
	MAPIFreeBuffer(statuses);
	debug() << "attendees needing secondary resolution:" << needingResolution.size();

	// Secondary resolution is to remove entries which have the the same 
	// email. But we must take care since we'll still have unresolved 
	// entries with empty email values (which would collide).
	QMap<QString, Attendee> uniqueResolvedAttendees;
	for (int i = 0; i < attendees.size(); i++) {
		Attendee &attendee = attendees[i];

		if (!attendee.email.isEmpty()) {
			// If we have duplicates, keep the one with the longer 
			// name in the hopes that it will be the more 
			// descriptive.
			QMap<QString, Attendee>::const_iterator i = uniqueResolvedAttendees.constFind(attendee.email);
			if (i != uniqueResolvedAttendees.constEnd()) {
				if (attendee.name.length() < i.value().name.length()) {
					continue;
				}
			}
			uniqueResolvedAttendees.insert(attendee.email, attendee);

			// If the name we just inserted matches one with an empty
			// email, we can kill the latter.
			for (int j = 0; j < needingResolution.size(); j++)
			{
				if (attendee.name == attendees[needingResolution.at(j)].name) {
					needingResolution.removeAt(j);
					j--;
				}
			}
		}
	}
	debug() << "attendees needing tertiary resolution:" << needingResolution.size();

	// Tertiary resolution.
	//
	// Add items needing more work. They should have an empty email. Just 
	// suck it up accepting the fact they'll be "duplicates" by key, but 
	// not by name. And delete any which have no name either.
	for (int i = 0; i < needingResolution.size(); i++)
	{
		Attendee &attendee = attendees[needingResolution.at(i)];

		if (attendee.name.isEmpty()) {
			// Grrrr...
			continue;
		}

		// Set the email to be the same as the name. It cannot be any
		// worse, right?
		attendee.email = attendee.name;
		uniqueResolvedAttendees.insertMulti(attendee.email, attendee);
	}

	// Finally, recreate the sanitised list.
	attendees.clear();
	foreach (Attendee attendee, uniqueResolvedAttendees) {
		attendees.append(attendee);
	}
	debug() << "attendees after resolution:" << attendees.size();
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
	if (!propertyWrite(PidTagSentRepresentingName, sender)) {
		return false;
	}
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

MapiConnector2::MapiConnector2() :
	MapiProfiles(),
	m_session(0)
{
	mapi_object_init(&m_store);
}

MapiConnector2::~MapiConnector2()
{
	if (m_session) {
		Logoff(&m_store);
	}
	mapi_object_release(&m_store);
}

QDebug MapiConnector2::debug() const
{
	static QString prefix = QString::fromAscii("MapiConnector2:");
	return TallocContext::debug(prefix);
}

bool MapiConnector2::defaultFolder(MapiDefaultFolder folderType, mapi_id_t *id)
{
	if (MAPI_E_SUCCESS != GetDefaultFolder(&m_store, id, folderType)) {
		error() << "cannot get default folder" << mapiError();
		return false;
	}
	return true;
}

QDebug MapiConnector2::error() const
{
	static QString prefix = QString::fromAscii("MapiConnector2:");
	return TallocContext::error(prefix);
}

bool MapiConnector2::fetchGAL(QList< GalMember >& list)
{
	struct SPropTagArray    *SPropTagArray;
	struct SRowSet      *SRowSet;
	enum MAPISTATUS     retval;
	uint32_t        i;
	uint32_t        count;
	uint8_t         ulFlags;
	uint32_t        rowsFetched = 0;
	uint32_t        totalRecs = 0;

	retval = GetGALTableCount(m_session, &totalRecs);
	if (retval != MAPI_E_SUCCESS) {
		return false;
	}
	qDebug() << "Total Number of entries in GAL:" << totalRecs;

	TALLOC_CTX *mem_ctx;
	mem_ctx = talloc_named(NULL, 0, "MapiGalConnector::fetchGAL");

	SPropTagArray = set_SPropTagArray(mem_ctx, 0xe,
						PR_INSTANCE_KEY,
						PR_ENTRYID,
						PR_DISPLAY_NAME_UNICODE,
						PR_EMAIL_ADDRESS_UNICODE,
						PR_DISPLAY_TYPE,
						PR_OBJECT_TYPE,
						PR_ADDRTYPE_UNICODE,
						PR_OFFICE_TELEPHONE_NUMBER_UNICODE,
						PR_OFFICE_LOCATION_UNICODE,
						PR_TITLE_UNICODE,
						PR_COMPANY_NAME_UNICODE,
						PR_ACCOUNT_UNICODE,
						PR_SMTP_ADDRESS_UNICODE,
						PR_SMTP_ADDRESS
 						);

	count = 0x7;
	ulFlags = TABLE_START;
	do {
		count += 0x2;
		retval = GetGALTable(m_session, SPropTagArray, &SRowSet, count, ulFlags);
		if ((!SRowSet) || (!(SRowSet->aRow))) {
			return false;
		}
		rowsFetched = SRowSet->cRows;
		if (rowsFetched) {
			for (i = 0; i < rowsFetched; i++) {
				GalMember data;

				for (unsigned int j = 0; j < SRowSet->aRow[i].cValues; ++j) {
					switch ( SRowSet->aRow[i].lpProps[j].ulPropTag ) {
						case PR_ENTRYID:
							data.id = QString::number( SRowSet->aRow[i].lpProps[j].value.l );
							break;
						case PR_DISPLAY_NAME_UNICODE:
							data.name = QString::fromUtf8( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						case PR_SMTP_ADDRESS_UNICODE:
							data.email = QString::fromUtf8( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						case PR_TITLE_UNICODE:
							data.title = QString::fromUtf8( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						case PR_COMPANY_NAME_UNICODE:
							data.organization = QString::fromUtf8( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						case PR_ACCOUNT_UNICODE:
							data.nick = QString::fromUtf8( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						default:
							break;
					}

					// If PR_ADDRTYPE_UNICODE == "EX" it's a NSPI address book entry (whatever that means?)
					// in that case the PR_EMAIL_ADDRESS_UNICODE contains a cryptic string. 
					// PR_SMTP_ADDRESS_UNICODE seams to work better

// 					TODO just for debugging
// 					QString tagStr;
// 					tagStr.sprintf("0x%X", SRowSet->aRow[i].lpProps[j].ulPropTag);
// 					qDebug() << tagStr << mapiValueToQString(&SRowSet->aRow[i].lpProps[j]);
				}

				if (data.id.isEmpty()) continue;

				qDebug() << "GLA:"<<data.id<<data.name<<data.email;
				list << data;
			}
		}
		ulFlags = TABLE_CUR;
		MAPIFreeBuffer(SRowSet);

// TODO for debugging
// 		break;
	} while (rowsFetched == count);

	MAPIFreeBuffer(SPropTagArray);

	return true;
}

bool MapiConnector2::login(QString profile)
{
	if (!init()) {
		return false;
	}
	if (m_session) {
		// already logged in...
		return true;
	}

	if (profile.isEmpty()) {
		// use the default profile if none was specified by the caller
		profile = defaultGet();
		if (profile.isEmpty()) {
			// there seams to be no default profile
			error() << "no default profile";
			return false;
		}
	}

	// Log on
	if (MAPI_E_SUCCESS != MapiLogonEx(m_context, &m_session, profile.toUtf8(), NULL)) {
		error() << "cannot logon using profile" << profile << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != OpenMsgStore(m_session, &m_store)) {
		error() << "cannot open message store" << mapiError();
		return false;
	}
	return true;
}

bool MapiConnector2::resolveNames(const char *names[], SPropTagArray *tags,
				  SRowSet **results, PropertyTagArray_r **statuses)
{
	if (MAPI_E_SUCCESS != ResolveNames(m_session, names, tags, results, statuses, MAPI_UNICODE)) {
		error() << "cannot resolve names" << mapiError();
		return false;
	}
	return true;
}

MapiFolder::MapiFolder(MapiConnector2 *connection, const char *tallocName, mapi_id_t id) :
	MapiObject(connection, tallocName, id)
{
	mapi_object_init(&m_contents);
	// A temporary name.
	name = QString::number(id, 36);
}

MapiFolder::~MapiFolder()
{
	mapi_object_release(&m_contents);
}

QDebug MapiFolder::debug() const
{
	static QString prefix = QString::fromAscii("MapiFolder: %1:");
	return MapiObject::debug(prefix.arg(m_id, ID_FORMAT));
}

QDebug MapiFolder::error() const
{
	static QString prefix = QString::fromAscii("MapiFolder: %1:");
	return MapiObject::error(prefix.arg(m_id, ID_FORMAT));
}

bool MapiFolder::childrenPull(QList<MapiFolder *> &children, const QString &filter)
{
	// Retrieve folder's folder table
	if (MAPI_E_SUCCESS != GetHierarchyTable(&m_object, &m_contents, 0, NULL)) {
		error() << "cannot get hierarchy table" << mapiError();
		return false;
	}

	// Create the MAPI table view
	SPropTagArray* tags = set_SPropTagArray(ctx(), 0x3, PidTagFolderId, PidTagDisplayName, PidTagContainerClass);
	if (!tags) {
		error() << "cannot set hierarchy table tags" << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != SetColumns(&m_contents, tags)) {
		error() << "cannot set hierarchy table columns" << mapiError();
		MAPIFreeBuffer(tags);
		return false;
	}
	MAPIFreeBuffer(tags);

	// Get current cursor position.
	uint32_t cursor;
	if (MAPI_E_SUCCESS != QueryPosition(&m_contents, NULL, &cursor)) {
		error() << "cannot query position" << mapiError();
		return false;
	}

	// Iterate through sets of rows.
	SRowSet rowset;
	while ((QueryRows(&m_contents, cursor, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS) && rowset.cRows) {
		for (unsigned i = 0; i < rowset.cRows; i++) {
			SRow &row = rowset.aRow[i];
			mapi_id_t fid;
			QString name;
			QString folderClass;

			for (unsigned j = 0; j < row.cValues; j++) {
				MapiProperty property(row.lpProps[j]); 

				// Note that the set of properties fetched here must be aligned
				// with those set above.
				switch (property.tag()) {
				case PidTagFolderId:
					fid = property.value().toULongLong(); 
					break;
				case PidTagDisplayName:
					name = property.value().toString(); 
					break;
				case PidTagContainerClass:
					folderClass = property.value().toString(); 
					break;
				default:
					//debug() << "ignoring folder property name:" << tagName(property.tag()) << property.value();
					break;
				}
			}
			if (!filter.isEmpty() && !folderClass.isEmpty() && !folderClass.startsWith(filter)) {
				debug() << "folder" << name << ", class" << folderClass << "does not match filter" << filter;
				continue;
			}

			// Add the entry to the output list!
			MapiFolder *data = new MapiFolder(m_connection, "MapiFolder::childrenPull", fid);
			data->name = name;
			children.append(data);
		}
	}
	return true;
}

bool MapiFolder::childrenPull(QList<MapiItem *> &children)
{
	// Retrieve folder's content table
	if (MAPI_E_SUCCESS != GetContentsTable(&m_object, &m_contents, 0, NULL)) {
		error() << "cannot get content table" << mapiError();
		return false;
	}

	// Create the MAPI table view
	SPropTagArray* tags = set_SPropTagArray(ctx(), 0x3, PidTagMid, PidTagConversationTopic, PidTagLastModificationTime);
	if (!tags) {
		error() << "cannot set content table tags" << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != SetColumns(&m_contents, tags)) {
		error() << "cannot set content table columns" << mapiError();
		MAPIFreeBuffer(tags);
		return false;
	}
	MAPIFreeBuffer(tags);

	// Get current cursor position.
	uint32_t cursor;
	if (MAPI_E_SUCCESS != QueryPosition(&m_contents, NULL, &cursor)) {
		error() << "cannot query position" << mapiError();
		return false;
	}

	// Iterate through sets of rows.
	SRowSet rowset;
	while ((QueryRows(&m_contents, cursor, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS) && rowset.cRows) {
		for (unsigned i = 0; i < rowset.cRows; i++) {
			SRow &row = rowset.aRow[i];
			mapi_id_t id;
			QString name;
			QDateTime modified;

			for (unsigned j = 0; j < row.cValues; j++) {
				MapiProperty property(row.lpProps[j]); 

				// Note that the set of properties fetched here must be aligned
				// with those set above.
				switch (property.tag()) {
				case PidTagMid:
					id = property.value().toULongLong(); 
					break;
				case PidTagConversationTopic:
					name = property.value().toString(); 
					break;
				case PidTagLastModificationTime:
					modified = property.value().toDateTime(); 
					break;
				default:
					//debug() << "ignoring item property name:" << tagName(property.tag()) << property.value();
					break;
				}
			}

			// Add the entry to the output list!
			MapiItem *data = new MapiItem(id, name, modified);
			children.append(data);
			//TODO Just for debugging (in case the content list ist very long)
			//if (i >= 10) break;
		}
	}
	return true;
}

bool MapiFolder::open()
{
	mapi_id_t id = m_id;

	// Get the toplevel folder id if needed.
	if (id == 0) {
		if (MAPI_E_SUCCESS != GetDefaultFolder(m_connection->d(), &id, olFolderTopInformationStore)) {
			error() << "cannot get default folder" << mapiError();
			return false;
		}
	}
	if (MAPI_E_SUCCESS != OpenFolder(m_connection->d(), id, &m_object)) {
		error() << "cannot open folder" << id << mapiError();
		return false;
	}
	return true;
}

MapiItem::MapiItem(mapi_id_t id, QString &name, QDateTime &modified) :
	m_id(id),
	m_name(name),
	m_modified(modified)
{
}

mapi_id_t MapiItem::id() const
{
	return m_id;
}

QString MapiItem::name() const
{
	return m_name;
}

QDateTime MapiItem::modified() const
{
	return m_modified;
}

MapiMessage::MapiMessage(MapiConnector2 *connection, const char *tallocName, mapi_id_t folderId, mapi_id_t id) :
	MapiObject(connection, tallocName, id),
	m_folderId(folderId)
{
	m_recipients.cRows = 0;
}

QDebug MapiMessage::debug() const
{
	static QString prefix = QString::fromAscii("MapiMessage: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT));
}

QDebug MapiMessage::error() const
{
	static QString prefix = QString::fromAscii("MapiMessage: %1/%2:");
	return MapiObject::error(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT));
}

mapi_id_t MapiMessage::folderId() const
{
	return m_folderId;
}

bool MapiMessage::open()
{
	if (MAPI_E_SUCCESS != OpenMessage(m_connection->d(), m_folderId, m_id, &m_object, 0x0)) {
		error() << "cannot open message, error:" << mapiError();
		return false;
	}
	return true;
}

Recipient MapiMessage::recipientAt(unsigned i) const
{
	Recipient result;

	if (i > m_recipients.cRows) {
		return Recipient();
	}

	struct SRow &recipient = m_recipients.aRow[i];
	for (unsigned j = 0; j < recipient.cValues; j++) {
		MapiProperty property(recipient.lpProps[j]);

		// Note that the set of properties fetched here must be aligned
		// with those fetched in MapiConnector::resolveNames().
		switch (property.tag()) {
		CASE_PREFER_UNICODE(PidTag7BitDisplayName, result.name, property.value().toString())
		CASE_PREFER_UNICODE(PidTagDisplayName, result.name, property.value().toString())
		CASE_PREFER_UNICODE(PidTagRecipientDisplayName, result.name, property.value().toString())
		CASE_PREFER_UNICODE(PidTagPrimarySmtpAddress, result.email, property.value().toString())
		case 0x6001001e:
			if (result.email.isEmpty()) {
				result.email = property.value().toString();
			}
			break;
		case 0x6001001f:
			result.email = property.value().toString(); 
			break;
		case PidTagRecipientTrackStatus:
			result.trackStatus = property.value().toInt();
			break;
		case PidTagRecipientFlags:
			result.flags = property.value().toInt();
			break;
		case PidTagRecipientType:
			result.type = property.value().toInt();
			break;
		case PidTagRecipientOrder:
			result.order = property.value().toInt();
			break;
		default:
			break;
		}
	}
	return result;
}

unsigned MapiMessage::recipientCount() const
{
	return m_recipients.cRows;
}

bool MapiMessage::recipientsPull()
{
	SPropTagArray propertyTagArray;

	if (MAPI_E_SUCCESS != GetRecipientTable(&m_object, &m_recipients, &propertyTagArray)) {
		error() << "cannot get recipient table:" << mapiError();
		return false;
	}
	return true;
}

MapiObject::MapiObject(MapiConnector2 *connection, const char *tallocName, mapi_id_t id) :
	TallocContext(tallocName),
	m_connection(connection),
	m_id(id),
	m_properties(0),
	m_propertyCount(0)
{
	mapi_object_init(&m_object);
}

MapiObject::~MapiObject()
{
	mapi_object_release(&m_object);
}

mapi_object_t *MapiObject::d() const
{
	return &m_object;
}

mapi_id_t MapiObject::id() const
{
	return m_id;
}

bool MapiObject::propertiesPush()
{
	if (MAPI_E_SUCCESS != SetProps(&m_object, MAPI_PROPS_SKIP_NAMEDID_CHECK, m_properties, m_propertyCount)) {
		error() << "cannot push:" << m_propertyCount << "properties:" << mapiError();
		return false;
	}
	m_properties = 0;
	m_propertyCount = 0;
	return true;
}

bool MapiObject::propertiesPull(QVector<int> &tags)
{
	struct SPropTagArray *tagArray = NULL;

	m_properties = 0;
	m_propertyCount = 0;
	foreach (int tag, tags) {
		if (!tagArray) {
			tagArray = set_SPropTagArray(ctx(), 1, (MAPITAGS)tag);
			if (!tagArray) {
				error() << "cannot allocate tags:" << mapiError();
				return false;
			}
		} else {
			if (MAPI_E_SUCCESS != SPropTagArray_add(ctx(), tagArray, (MAPITAGS)tag)) {
				error() << "cannot extend tags:" << mapiError();
				MAPIFreeBuffer(tagArray);
				return false;
			}
		}
	}
	if (MAPI_E_SUCCESS != GetProps(&m_object, MAPI_UNICODE, tagArray, &m_properties, &m_propertyCount)) {
		error() << "cannot pull properties:" << mapiError();
		MAPIFreeBuffer(tagArray);
		return false;
	}
	MAPIFreeBuffer(tagArray);
	return true;
}

bool MapiObject::propertiesPull()
{
	struct mapi_SPropValue_array mapiProperties;

	m_properties = 0;
	m_propertyCount = 0;
	if (MAPI_E_SUCCESS != GetPropsAll(&m_object, MAPI_UNICODE, &mapiProperties)) {
		error() << "cannot pull all properties:" << mapiError();
		return false;
	}

	// Copy results from MAPI array to our array.
	m_properties = talloc_array(ctx(), SPropValue, mapiProperties.cValues);
	if (m_properties) {
		for (m_propertyCount = 0; m_propertyCount < mapiProperties.cValues; m_propertyCount++) {
			cast_SPropValue(ctx(), &mapiProperties.lpProps[m_propertyCount], 
					&m_properties[m_propertyCount]);
		}
	} else {
		error() << "cannot copy properties:" << mapiError();
	}
	return true;
}

QVariant MapiObject::property(int tag) const
{
	return propertyAt(propertyFind(tag));
}

QVariant MapiObject::propertyAt(unsigned i) const
{
	if (!m_propertyCount || ((m_propertyCount - 1) < i)) {
		return QVariant();
	}
	return MapiProperty(m_properties[i]).value();
}

unsigned MapiObject::propertyCount() const
{
	return m_propertyCount;
}

unsigned MapiObject::propertyFind(int tag) const
{
	for (unsigned i = 0; i < m_propertyCount; i++) {
		if (m_properties[i].ulPropTag == tag) {
			return i;
		}
	}
	return UINT_MAX;
}

QString MapiObject::propertyString(unsigned i) const
{
	if (!m_propertyCount || ((m_propertyCount - 1) < i)) {
		return QString();
	}
	return MapiProperty(m_properties[i]).toString();
}

bool MapiObject::propertyWrite(int tag, void *data, bool idempotent)
{
	// If the assignment is idempotent, if an instance of the 
	// property exists, it will be overwritten.
	if (idempotent) {
		for (unsigned i = 0; i < m_propertyCount; i++) {
			if (m_properties[i].ulPropTag == tag) {
				bool ok = set_SPropValue_proptag(&m_properties[i], (MAPITAGS)tag, data);
				if (!ok) {
					error() << "cannot overwrite tag:" << tagName(tag) << "value:" << data;
				}
				return ok;
			}
		}
	}

	// Add a new entry to the array.
	m_properties = add_SPropValue(ctx(), m_properties, &m_propertyCount, (MAPITAGS)tag, data);
	if (!m_properties) {
		error() << "cannot write tag:" << tagName(tag) << "value:" << data;
		return false;
	}
	return true;
}

bool MapiObject::propertyWrite(int tag, int data, bool idempotent)
{
	return propertyWrite(tag, &data, idempotent);
}

bool MapiObject::propertyWrite(int tag, QString &data, bool idempotent)
{
	char *copy = talloc_strdup(ctx(), data.toUtf8().data());

	if (!copy) {
		error() << "cannot talloc:" << data;
		return false;
	}

	return propertyWrite(tag, copy, idempotent);
}

bool MapiObject::propertyWrite(int tag, QDateTime &data, bool idempotent)
{
	FILETIME *copy = talloc(ctx(), FILETIME);

	if (!copy) {
		error() << "cannot talloc:" << data;
		return false;
	}

	// As per http://support.citrix.com/article/CTX109645.
	time_t unixTime = data.toTime_t();
	NTTIME ntTime = (unixTime + 11644473600L) * 10000000;
	copy->dwHighDateTime = ntTime >> 32;
	copy->dwLowDateTime = ntTime;
	return propertyWrite(tag, copy, idempotent);
}

QString MapiObject::tagAt(unsigned i) const
{
	if (!m_propertyCount || ((m_propertyCount - 1) < i)) {
		return QString();
	}
	return tagName(m_properties[i].ulPropTag);
}

QString MapiObject::tagName(int tag) const
{
	const char *str = get_proptag_name(tag);

	if (str) {
		return QString::fromLatin1(str);
	} else {
		struct MAPINAMEID *names;
		uint16_t count;
		int safeTag = (tag & 0xFFFF0000) | PT_NULL;

		/*
			* Try a lookup.
			*/
		if (MAPI_E_SUCCESS != GetNamesFromIDs(&m_object, (MAPITAGS)safeTag, &count, &names)) {
			return QString::fromLatin1("Pid0x%1").arg(tag, 0, 16);
		} else {
			QByteArray strs;

			/*
				* Oh dear, a lookup can return multiple 
				* names...
				*/
			for (unsigned i = 0; i < count; i++) {
				if (i) {
					strs.append(',');
				}
				if (MNID_STRING == names[i].ulKind) {
					strs.append(&names[i].kind.lpwstr.Name[0], names[i].kind.lpwstr.NameSize);
				} else {
					strs.append(QString::fromLatin1("Id0x%1:%2").arg((unsigned)tag, 0, 16).
							arg((unsigned)names[i].kind.lid, 0, 16).toLatin1());
				}
			}
			return QString::fromLatin1(strs);
		}
	}
}

MapiProfiles::MapiProfiles() :
	TallocContext("MapiProfiles::MapiProfiles"),
	m_context(0),
	m_initialised(false)
{
}

MapiProfiles::~MapiProfiles()
{
	if (m_context) {
		MAPIUninitialize(m_context);
	}
}

bool MapiProfiles::add(QString profile, QString username, QString password, QString domain, QString server)
{
	if (!init()) {
		return false;
	}

	const char *profile8 = profile.toUtf8();

	if (MAPI_E_SUCCESS != CreateProfile(m_context, profile8, username.toUtf8(), password.toUtf8(), 0)) {
		error() << "cannot create profile:" << mapiError();
		return false;
	}

	// TODO get workstation as parameter (was is it needed for anyway?)
	char hostname[256] = {};
	gethostname(&hostname[0], sizeof(hostname) - 1);
	hostname[sizeof(hostname) - 1] = 0;
	QString workstation = QString::fromLatin1(hostname);

	if (!addAttribute(profile8, "binding", server)) {
		error() << "cannot add binding:" << server << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "workstation", workstation)) {
		error() << "cannot add workstation:" << workstation << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "domain", domain)) {
		error() << "cannot add domain:" << domain << mapiError();
		return false;
	}
// What is seal for? Seams to have something to do with Exchange 2010
// 	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "seal", (seal == true) ? "true" : "false");

// TODO Get langage from parameter if needed
// 	const char* locale = (const char *) (language) ? mapi_get_locale_from_language(language) : mapi_get_system_locale();
	const char *locale = mapi_get_system_locale();
	if (!locale) {
		error() << "cannot find system locale:" << mapiError();
		return false;
	}

	uint32_t cpid = mapi_get_cpid_from_locale(locale);
	uint32_t lcid = mapi_get_lcid_from_locale(locale);
	if (!cpid || !lcid) {
		error() << "invalid Locale supplied or unknown system locale" << locale << ", deleting profile..." << mapiError();
		if (!remove(profile)) {
			return false;
		}
		return false;
	}

	if (!addAttribute(profile8, "codepage", QString::number(cpid))) {
		error() << "cannot add codepage:" << cpid << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "language", QString::number(lcid))) {
		error() << "cannot language:" << lcid << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "method", QString::number(lcid))) {
		error() << "cannot method:" << lcid << mapiError();
		return false;
	}

	struct mapi_session *session = NULL;
	if (MAPI_E_SUCCESS != MapiLogonProvider(m_context, &session, profile8, password.toUtf8(), PROVIDER_ID_NSPI)) {
		error() << "cannot get logon provider, deleting profile..." << mapiError();
		if (!remove(profile)) {
			return false;
		}
		return false;
	}

	int retval = ProcessNetworkProfile(session, username.toUtf8().constData(), profileSelectCallback, NULL);
	if (retval != MAPI_E_SUCCESS && retval != 0x1) {
		error() << "cannot process network profile, deleting profile..." << mapiError();
		if (!remove(profile)) {
			return false;
		}
		return false;
	}
	return true;
}

bool MapiProfiles::addAttribute(const char *profile, const char *attribute, QString value) 
{
	if (MAPI_E_SUCCESS != mapi_profile_add_string_attr(m_context, profile, attribute, value.toUtf8())) {
		return false;
	}
	return true;
}

QDebug MapiProfiles::debug() const
{
	static QString prefix = QString::fromAscii("MapiProfiles:");
	return TallocContext::debug(prefix);
}

QString MapiProfiles::defaultGet()
{
	if (!init()) {
		return QString();
	}

	char *profname;
	if (MAPI_E_SUCCESS != GetDefaultProfile(m_context, &profname)) {
		error() << "cannot get default profile:" << mapiError();
		return QString();
	}
	return QString::fromLocal8Bit(profname);
}

bool MapiProfiles::defaultSet(QString profile)
{
	if (!init()) {
		return false;
	}

	if (MAPI_E_SUCCESS != SetDefaultProfile(m_context, profile.toUtf8())) {
		error() << "cannot set default profile:" << profile << mapiError();
		return false;
	}
	return true;
}

QDebug MapiProfiles::error() const
{
	static QString prefix = QString::fromAscii("MapiProfiles:");
	return TallocContext::error(prefix);
}

bool MapiProfiles::init()
{
	if (m_initialised) {
		return true;
	}
	QString profilePath(QDir::home().path());
	profilePath.append(QLatin1String("/.openchange/"));
	
	QString profileFile(profilePath);
	profileFile.append(QString::fromLatin1("profiles.ldb"));

	// Check if the store exists.
	QDir path(profilePath);
	if (!path.exists()) {
		if (path.mkpath(profilePath)) {
			error() << "cannot make profile path:" << profilePath;
			return false;
		}
	}
	if (!QFile::exists(profileFile)) {
		if (MAPI_E_SUCCESS != CreateProfileStore(profileFile.toUtf8(), mapi_profile_get_ldif_path())) {
			error() << "cannot create profile store:" << profileFile << mapiError();
			return false;
		}
	}
	if (MAPI_E_SUCCESS != MAPIInitialize(&m_context, profileFile.toLatin1())) {
		error() << "cannot init profile store:" << profileFile << mapiError();
		return false;
	}
	m_initialised = true;
	return true;
}

QStringList MapiProfiles::list()
{
	if (!init()) {
		return QStringList();
	}

	struct SRowSet proftable;

	if (MAPI_E_SUCCESS != GetProfileTable(m_context, &proftable)) {
		error() << "cannot get profile table" << mapiError();
		return QStringList();
	}

	// debug() << "Profiles in the database:" << proftable.cRows;
	QStringList profiles;
	for (unsigned count = 0; count < proftable.cRows; count++) {
		const char *name = proftable.aRow[count].lpProps[0].value.lpszA;
		uint32_t dflt = proftable.aRow[count].lpProps[1].value.l;

		profiles.append(QString::fromLocal8Bit(name));
		if (dflt) {
			debug() << "default profile:" << name;
		}
	}
	return profiles;
}

bool MapiProfiles::remove(QString profile)
{
	if (!init()) {
		return false;
	}

	if (MAPI_E_SUCCESS != DeleteProfile(m_context, profile.toUtf8())) {
		error() << "cannot delete profile:" << profile << mapiError();
		return false;
	}
	return true;
}

MapiProperty::MapiProperty(SPropValue &property) :
	m_property(property)
{
}

int MapiProperty::tag() const
{
	return m_property.ulPropTag;
}

/**
 * Get the string equivalent of a property, e.g. for display purposes.
 * We take care to hex-ify GUIDs and other byte arrays, and lists of
 * the same.
 */
QString MapiProperty::toString() const
{
	// Use the default stringification whenever we can.
	QVariant tmp(value());
	switch (tmp.type()) {
	case QVariant::ByteArray:
		// Convert to a hex string.
		return QString::fromLatin1(tmp.toByteArray().toHex());
	case QVariant::List:
	{
		QList<QVariant> list(tmp.toList());
		QList<QVariant>::iterator i;
		QStringList result;

		for (i = list.begin(); i != list.end(); ++i) {
			switch ((*i).type()) {
			case QVariant::ByteArray:
				// Convert to a hex string.
				result.append(QString::fromLatin1((*i).toByteArray().toHex()));
				break;
			default:
				result.append((*i).toString());
				break;
			}
		}
		return result.join(QString::fromLatin1(","));
	}
	default:
		return tmp.toString();
	}
}

/**
 * Get the value of the property in a nice typesafe wrapper.
 */
inline
QVariant MapiProperty::value() const
{
	switch (m_property.ulPropTag & 0xFFFF) {
	case PT_SHORT:
		return m_property.value.i;
	case PT_LONG:
		return m_property.value.l;
	case PT_FLOAT:
		return (float)m_property.value.l;
	case PT_DOUBLE:
		return (double)m_property.value.dbl;
	case PT_BOOLEAN:
		return m_property.value.b;
	case PT_I8:
		return (qlonglong)m_property.value.d;
	case PT_STRING8:
		return QString::fromLocal8Bit(m_property.value.lpszA);
	case PT_BINARY:
	case PT_SVREID:
		return QByteArray((char *)m_property.value.bin.lpb, m_property.value.bin.cb);
	case PT_UNICODE:
		return QString::fromUtf8(m_property.value.lpszW);
	case PT_CLSID:
		return QByteArray((char *)&m_property.value.lpguid->ab[0], 
					sizeof(m_property.value.lpguid->ab));
	case PT_SYSTIME:
		return convertSysTime(m_property.value.ft);
	case PT_ERROR:
		return (unsigned)m_property.value.err;
	case PT_MV_SHORT:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVi.cValues; i++) {
			ret.append(m_property.value.MVi.lpi[i]);
		}
		return ret;
	}
	case PT_MV_LONG:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVl.cValues; i++) {
			ret.append(m_property.value.MVl.lpl[i]);
		}
		return ret;
	}
	case PT_MV_FLOAT:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVl.cValues; i++) {
			ret.append((float)m_property.value.MVl.lpl[i]);
		}
		return ret;
	}
	case PT_MV_STRING8:
	{
		QStringList ret;

		for (unsigned i = 0; i < m_property.value.MVszA.cValues; i++) {
			ret.append(QString::fromLocal8Bit(m_property.value.MVszA.lppszA[i]));
		}
		return ret;
	}
	case PT_MV_BINARY:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVszW.cValues; i++) {
			ret.append(QByteArray((char *)m_property.value.MVbin.lpbin[i].lpb, 
						m_property.value.MVbin.lpbin[i].cb));
		}
		return ret;
	}
	case PT_MV_CLSID:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVguid.cValues; i++) {
			ret.append(QByteArray((char *)&m_property.value.MVguid.lpguid[i]->ab[0], 
						sizeof(m_property.value.MVguid.lpguid[i]->ab)));
		}
		return ret;
	}
	case PT_MV_UNICODE:
	{
		QStringList ret;

		for (unsigned i = 0; i < m_property.value.MVszW.cValues; i++) {
			ret.append(QString::fromUtf8(m_property.value.MVszW.lppszW[i]));
		}
		return ret;
	}
	case PT_MV_SYSTIME:
	{
		QList<QVariant> ret;

		for (unsigned i = 0; i < m_property.value.MVft.cValues; i++) {
			ret.append(convertSysTime(m_property.value.MVft.lpft[i]));
		}
		return ret;
	}
	case PT_NULL:
		return m_property.value.null;
	case PT_OBJECT:
		return m_property.value.object;
	default:
		return QString::fromLatin1("PT_0x%1").arg(m_property.ulPropTag & 0xFFFF, 0, 16);
	}
}

bool MapiRecurrencyPattern::setData(RecurrencePattern* pattern)
{
	if (pattern->RecurFrequency == RecurFrequency_Daily && pattern->PatternType == PatternType_Day) {
		this->mRecurrencyType = Daily;
		this->mPeriod = (pattern->Period / 60) / 24;
	} 
	else if (pattern->RecurFrequency == RecurFrequency_Daily && pattern->PatternType == PatternType_Week) {
		this->mRecurrencyType = Every_Weekday;
		this->mPeriod = 0;

		QBitArray bitArray(7, true); // everyday ...
		bitArray.setBit(5, false); // ... except saturday ..
		bitArray.setBit(6, false); // ... except sunday
		this->mDays = bitArray;

		this->mFirstDOW = convertDayOfWeek(pattern->FirstDOW);
	}
	else if (pattern->RecurFrequency == RecurFrequency_Weekly && pattern->PatternType == PatternType_Week) {
		this->mRecurrencyType = Weekly;
		this->mPeriod = pattern->Period;
		this->mDays = getRecurrenceDays(pattern->PatternTypeSpecific.WeekRecurrencePattern);
		this->mFirstDOW = convertDayOfWeek(pattern->FirstDOW);
	}
	else if (pattern->RecurFrequency == RecurFrequency_Monthly) {
		this->mRecurrencyType = Monthly;
		this->mPeriod = pattern->Period;
	}
	else if (pattern->RecurFrequency == RecurFrequency_Yearly) {
		this->mRecurrencyType = Yearly;
		this->mPeriod = pattern->Period;
	} else {
		qCritical() << "unsupported frequency:"<<pattern->RecurFrequency;
		return false;
	}

	this->mStartDate = convertExchangeTimes(pattern->StartDate);
	this->mOccurrenceCount = 0;
	switch (pattern->EndType) {
		case END_AFTER_DATE:
			this->mEndType = Date;
			this->mEndDate = convertExchangeTimes(pattern->EndDate);
			break;
		case END_AFTER_N_OCCURRENCES:
			this->mEndType = Count;
			this->mOccurrenceCount = pattern->OccurrenceCount;
			break;
		case END_NEVER_END:
			this->mEndType = Never;
			break;
		default:
			qCritical() << "unsupported endtype:"<<pattern->EndType;
			return false;
	}

	this->setRecurring(true);

	return true;
}

QBitArray MapiRecurrencyPattern::getRecurrenceDays(const uint32_t exchangeDays)
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

int MapiRecurrencyPattern::convertDayOfWeek(const uint32_t exchangeDayOfWeek)
{
	int retVal = exchangeDayOfWeek;
	if (retVal == 0) {
		// Exchange-Sunday(0) mapped to KCal-Sunday(7)
		retVal = 7;
	}
	return retVal;
}

QDateTime MapiRecurrencyPattern::convertExchangeTimes(const uint32_t exchangeMinutes)
{
	// exchange stores the recurrency times as minutes since 1.1.1601
	QDateTime calc(QDate(1601, 1, 1));
	int days = exchangeMinutes / 60 / 24;
	int secs = exchangeMinutes - (days*24*60);
	return calc.addDays(days).addSecs(secs);
}

TallocContext::TallocContext(const char *name)
{
	m_ctx = talloc_named(NULL, 0, "%s", name);
}

TallocContext::~TallocContext()
{
	talloc_free(m_ctx);
}

TALLOC_CTX *TallocContext::ctx()
{
	return m_ctx;
}

QDebug TallocContext::debug(const QString &caller) const
{
	static QString prefix = QString::fromAscii("%1.%2");
	QString talloc = QString::fromAscii(talloc_get_name(m_ctx));
	return qDebug() << prefix.arg(talloc).arg(caller);
}

QDebug TallocContext::error(const QString &caller) const
{
	static QString prefix = QString::fromAscii("%1.%2");
	QString talloc = QString::fromAscii(talloc_get_name(m_ctx));
	return qCritical() << prefix.arg(talloc).arg(caller);
}

MapiNote::MapiNote(MapiConnector2 *connector, const char *tallocName, mapi_id_t folderId, mapi_id_t id) :
	MapiMessage(connector, tallocName, folderId, id)
{
}

QDebug MapiNote::debug() const
{
	static QString prefix = QString::fromAscii("MapiNote: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT)) /*<< title*/;
}

QDebug MapiNote::error() const
{
	static QString prefix = QString::fromAscii("MapiNote: %1/%2");
	return MapiObject::error(prefix.arg(m_folderId, ID_FORMAT).arg(m_id, ID_FORMAT)) /*<< title*/;
}

bool MapiNote::propertiesPull()
{
#if (DEBUG_NOTE_PROPERTIES)
	if (!MapiMessage::propertiesPull()) {
		return false;
	}
#else
	QVector<int> readTags;
	readTags.append(PidTagMessageClass);
	readTags.append(PidTagDisplayTo);
	readTags.append(PidTagConversationTopic);
	readTags.append(PidTagBody);
	readTags.append(PidTagLastModificationTime);
	readTags.append(PidTagCreationTime);
	if (!MapiMessage::propertiesPull(readTags)) {
		return false;
	}
#endif

	QStringList displayTo;

	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		switch (property.tag()) {
			// Sanity check the message class.
			if (QLatin1String("IPM.Note") != property.value().toString()) {
				// this one is not an appointment
				return false;
			}
			case PidTagDisplayTo:
			// Step 2. Add the DisplayTo items.
			//
			// Astonishingly, when we call recipientsPull() later,
			// the results can contain entries with missing name (!)
			// and email values. Reading this property is a pathetic
			// workaround.
			displayTo = property.value().toString().split(QChar::fromAscii(';'));
			foreach (QString name, displayTo) {
			}
			break;
		case PidTagConversationTopic:
			title = property.value().toString();
			break;
		case PidTagBody:
			text = property.value().toString();
			break;
		case PidTagCreationTime:
			created = property.value().toDateTime();
			break;
//		CASE_PREFER_UNICODE(PidTagSentRepresentingEmailAddress, sentRepresenting.email, property.value().toString())
//		CASE_PREFER_UNICODE(PidTagSentRepresentingName, sentRepresenting.name, property.value().toString())
//		CASE_PREFER_UNICODE(PidTagSentRepresentingSimpleDisplayName, sentRepresenting.name, property.value().toString())
//		CASE_PREFER_UNICODE(PidTagOriginalSentRepresentingEmailAddress, originalSentRepresenting.email, property.value().toString())
//		CASE_PREFER_UNICODE(PidTagOriginalSentRepresentingName, originalSentRepresenting.name, property.value().toString())
		default:
#if (DEBUG_NOTE_PROPERTIES)
			debug() << "ignoring note property:" << tagName(property.tag()) << property.value();
#endif
			break;
		}
	}
	return true;
}

bool MapiNote::propertiesPush()
{
	// Overwrite all the fields we know about.
	if (!propertyWrite(PidTagConversationTopic, title)) {
		return false;
	}
	if (!propertyWrite(PidTagBody, text)) {
		return false;
	}
	if (!propertyWrite(PidTagCreationTime, created)) {
		return false;
	}
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
