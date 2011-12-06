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
#include <QVariant>

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
	qDebug() << "Found more than 1 matching users -> cancel";

	//  TODO Some sort of handling would be needed here
	return rowset->cRows;
}

MapiAppointment::MapiAppointment(MapiConnector2 *connector, const char *tallocName, mapi_id_t id) :
	MapiMessage(connector, tallocName, id)
{
}

bool MapiAppointment::debugRecurrencyPattern(RecurrencePattern *pattern)
{
	// do the actual work
	qDebug() << "-- Recurrency debug output [BEGIN] --";
	switch (pattern->RecurFrequency) {
	case RecurFrequency_Daily:
		qDebug() << "Fequency: daily";
		break;
	case RecurFrequency_Weekly:
		qDebug() << "Fequency: weekly";
		break;
	case RecurFrequency_Monthly:
		qDebug() << "Fequency: monthly";
		break;
	case RecurFrequency_Yearly:
		qDebug() << "Fequency: yearly";
		break;
	default:
		qDebug() << "unsupported frequency:"<<pattern->RecurFrequency;
		return false;
	}

	switch (pattern->PatternType) {
	case PatternType_Day:
		qDebug() << "PatternType: day";
		break;
	case PatternType_Week:
		qDebug() << "PatternType: week";
		break;
	case PatternType_Month:
		qDebug() << "PatternType: month";
		break;
	default:
		qDebug() << "unsupported patterntype:"<<pattern->PatternType;
		return false;
	}

	qDebug() << "Calendar:" << pattern->CalendarType;
	qDebug() << "FirstDateTime:" << pattern->FirstDateTime;
	qDebug() << "Period:" << pattern->Period;
	if (pattern->PatternType == PatternType_Month) {
		qDebug() << "PatternTypeSpecific:" << pattern->PatternTypeSpecific.Day;
	} else if (pattern->PatternType == PatternType_Week) {
		qDebug() << "PatternTypeSpecific:" << QString::number(pattern->PatternTypeSpecific.WeekRecurrencePattern, 2);
	}

	switch (pattern->EndType) {
		case END_AFTER_DATE:
			qDebug() << "EndType: after date";
			break;
		case END_AFTER_N_OCCURRENCES:
			qDebug() << "EndType: after occurenc count";
			break;
		case END_NEVER_END:
			qDebug() << "EndType: never";
			break;
		default:
			qDebug() << "unsupported endtype:"<<pattern->EndType;
			return false;
	}
	qDebug() << "OccurencCount:" << pattern->OccurrenceCount;
	qDebug() << "FirstDOW:" << pattern->FirstDOW;
	qDebug() << "Start:" << pattern->StartDate;
	qDebug() << "End:" << pattern->EndDate;
	qDebug() << "-- Recurrency debug output [END] --";

	return true;
}

bool MapiAppointment::getAttendees()
{
	if (!MapiMessage::recipientsPull()) {
		return false;
	}

	// Get the attendees, and find any that need resolution.
	QList<unsigned> needingResolution;
	for (unsigned i = 0; i < recipientCount(); i++) {
// TODO for debugging
//		QString tagStr;
//		tagStr.sprintf("0x%X", recipients.aRow[i].lpProps[j].ulPropTag);
//		qDebug() << tagStr << mapiValueToQString(&recipients.aRow[i].lpProps[j]);
		Attendee attendee(recipientAt(i));
		if (attendee.email.isEmpty() ||
			!attendee.email.contains(QString::fromAscii("@"))) {
			needingResolution << i;
			qCritical() << "needingResolution:" << attendee.name << attendee.email;
		}
		attendees << attendee;
	}

	if (needingResolution.size() == 0) {
		return true;
	}

	// Fill an array with the names we need to resolve.
	const char *names[needingResolution.size() + 1];
	unsigned j = 0;
	foreach (unsigned i, needingResolution) {
		Attendee &attendee = attendees[i];
		names[j] = talloc_strdup(ctx(), attendee.name.toStdString().c_str());
		j++;
	}
	names[j] = 0;

	// Resolve the lot.
	struct PropertyTagArray_r *statuses = NULL;
	struct SPropTagArray *tags = NULL;
	struct SRowSet *results = NULL;

	tags = set_SPropTagArray(ctx(), 0x9, 
				 PR_7BIT_DISPLAY_NAME,         PR_DISPLAY_NAME,         PR_RECIPIENT_DISPLAY_NAME, 
				 PR_7BIT_DISPLAY_NAME_UNICODE, PR_DISPLAY_NAME_UNICODE, PR_RECIPIENT_DISPLAY_NAME_UNICODE, 
				 PR_SMTP_ADDRESS,
				 PR_SMTP_ADDRESS_UNICODE,
				 0x6001001f);
	if (!m_connection->resolveNames(names, tags, &results, &statuses)) {
		return false;
	}
	if (results) {
		for (unsigned i = 0; i < results->cRows; i++) {
			if (MAPI_UNRESOLVED != statuses->aulPropTag[i]) {
				struct SRow &recipient = results->aRow[i];
				QString name, email;
				for (unsigned j = 0; j < recipient.cValues; j++) {
					MapiProperty property(recipient.lpProps[j]);

					// Note that the set of properties fetched here must be aligned
					// with those fetched in MapiMessage::recipientAt(), as well
					// as the array above.
					switch (property.tag()) {
					case PR_7BIT_DISPLAY_NAME:
					case PR_DISPLAY_NAME:
					case PR_RECIPIENT_DISPLAY_NAME:
						// We'd prefer a Unicode answer if there is one!
						if (name.isEmpty()) {
							name = property.value().toString(); 
						}
						break;
					case PR_7BIT_DISPLAY_NAME_UNICODE:
					case PR_DISPLAY_NAME_UNICODE:
					case PR_RECIPIENT_DISPLAY_NAME_UNICODE:
						name = property.value().toString(); 
						break;
					case PR_SMTP_ADDRESS:
						// We'd prefer a Unicode answer if there is one!
						if (email.isEmpty()) {
							email = property.value().toString(); 
						}
						break;
					case PR_SMTP_ADDRESS_UNICODE:
					case 0x6001001f:
						email = property.value().toString(); 
						break;
					default:
						break;
					}
				}

				// A resolved value is better than an unresolved one.
				Attendee &attendee = attendees[needingResolution.at(i)];
				if (!name.isEmpty()) {
					attendee.name = name;
				}
				if (!email.isEmpty()) {
					attendee.email = email;
				}
			}
		}
	}
	MAPIFreeBuffer(results);
	MAPIFreeBuffer(statuses);

	// Time for one final fallback, based on the use of the name.
	foreach (unsigned i, needingResolution) {
		Attendee &attendee = attendees[i];
		if (attendee.email.isEmpty()) {
			attendee.email = attendee.name;
		}
	}
	foreach (const Attendee &attendee, attendees) {
		qCritical () << "Attendee" << attendee.name << attendee.email;
	}
	return true;
}

bool MapiAppointment::open(mapi_id_t folderId)
{
	if (!MapiMessage::open(folderId)) {
		return false;
	}

	// Sanity check the message class.
	QVector<int> readTags;
	readTags.append(PidTagMessageClass);
	if (!MapiMessage::propertiesPull(readTags)) {
		return false;
	}
	QString messageClass = property(PidTagMessageClass).toString();
	if (QString::fromLatin1("IPM.Appointment").compare(messageClass)) {
		// this one is not an appointment
		return false;
	}
	fid = QString::number(folderId);
	id = QString::number(m_id);
	reminderActive = false;
	return true;
}

bool MapiAppointment::propertiesPull()
{
	if (!MapiMessage::propertiesPull()) {
		return false;
	}

	unsigned recurrenceType = 0;
	RecurrencePattern *pattern = 0;
	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		// Note that the set of properties fetched here must be aligned
		// with those set above.
		switch (property.tag()) {
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
		case PidTagSentRepresentingName:
			sender = property.value().toString();
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
			//qCritical() << "ignoring appointment property name:" << tagName(property.tag()) << property.value();
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
			qCritical() << "missing recurrencePattern in message" << m_id;
		}
		return false;
	}

	if (!getAttendees()) {
		return false;
	}
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
			qDebug() << "missing recurrencePattern in message"<<messageID<<"in folder"<<folderID;
			}
	}

	getAttendees(obj_message, QString::fromLocal8Bit(toAttendeesString), data);
#endif
	if (!MapiMessage::propertiesPush()) {
		return false;
	}
#if 0
	qCritical() << "************  OpenFolder";
	if (!OpenFolder(&m_store, folder.id(), folder.d())) {
		qCritical() << "cannot open folder" << folderID
			<< ", error:" << mapiError();
		return false;
        }
	qCritical() << "************  SaveChangesMessage";
	if (!SaveChangesMessage(folder.d(), message.d(), KeepOpenReadWrite)) {
		qCritical() << "cannot save message" << messageID << "in folder" << folderID
			<< ", error:" << mapiError();
		return false;
	}
#endif
	qCritical() << "************  SubmitMessage";
	if (MAPI_E_SUCCESS != SubmitMessage(&m_object)) {
		qCritical() << "cannot submit message" << id << "in folder" << fid
			<< ", error:" << mapiError();
		return false;
	}
	struct mapi_SPropValue_array replyProperties;
	qCritical() << "************  TransportSend";
	if (MAPI_E_SUCCESS != TransportSend(&m_object, &replyProperties)) {
		qDebug() << "cannot send message" << id << "in folder" << fid
			<< ", error:" << mapiError();
		return false;
	}
	return true;
}

MapiFolder::MapiFolder(MapiConnector2 *connection, const char *tallocName, mapi_id_t id) :
	MapiObject(connection, tallocName, id)
{
	mapi_object_init(&m_contents);
	// A temporary name.
	name = this->id();
}

MapiFolder::~MapiFolder()
{
	mapi_object_release(&m_contents);
}

QString MapiFolder::id() const
{
	return QString::number(m_id);
}

bool MapiFolder::childrenPull(QList<MapiFolder *> &children, const QString &filter)
{
	// Retrieve folder's folder table
	if (MAPI_E_SUCCESS != GetHierarchyTable(&m_object, &m_contents, 0, NULL)) {
		qCritical() << "cannot get hierarchy table" << mapiError();
		return false;
	}

	// Create the MAPI table view
	SPropTagArray* tags = set_SPropTagArray(ctx(), 0x3, PR_FID, PR_DISPLAY_NAME, PR_CONTAINER_CLASS);
	if (!tags) {
		qCritical() << "cannot set hierarchy table tags" << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != SetColumns(&m_contents, tags)) {
		qCritical() << "cannot set hierarchy table columns" << mapiError();
		MAPIFreeBuffer(tags);
		return false;
	}
	MAPIFreeBuffer(tags);

	// Get current cursor position.
	uint32_t cursor;
	if (MAPI_E_SUCCESS != QueryPosition(&m_contents, NULL, &cursor)) {
		qCritical() << "cannot query position" << mapiError();
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
				case PR_FID:
					fid = property.value().toULongLong(); 
					break;
				case PR_DISPLAY_NAME:
					name = property.value().toString(); 
					break;
				case PR_CONTAINER_CLASS:
					folderClass = property.value().toString(); 
					break;
				default:
					//qCritical() << "ignoring folder property name:" << tagName(property.tag()) << property.value();
					break;
				}
			}
			if (!filter.isEmpty() && !(folderClass).startsWith(filter)) {
				qCritical() << "Folder" << name << "does not match filter" << filter;
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
		qCritical() << "cannot get content table" << mapiError();
		return false;
	}

	// Create the MAPI table view
	SPropTagArray* tags = set_SPropTagArray(ctx(), 0x4, PR_FID, PR_MID, PR_CONVERSATION_TOPIC, PR_LAST_MODIFICATION_TIME);
	if (!tags) {
		qCritical() << "cannot set content table tags" << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != SetColumns(&m_contents, tags)) {
		qCritical() << "cannot set content table columns" << mapiError();
		MAPIFreeBuffer(tags);
		return false;
	}
	MAPIFreeBuffer(tags);

	// Get current cursor position.
	uint32_t cursor;
	if (MAPI_E_SUCCESS != QueryPosition(&m_contents, NULL, &cursor)) {
		qCritical() << "cannot query position" << mapiError();
		return false;
	}

	// Iterate through sets of rows.
	SRowSet rowset;
	while ((QueryRows(&m_contents, cursor, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS) && rowset.cRows) {
		for (unsigned i = 0; i < rowset.cRows; i++) {
			SRow &row = rowset.aRow[i];
			MapiItem *data = new MapiItem();

			for (unsigned j = 0; j < row.cValues; j++) {
				MapiProperty property(row.lpProps[j]); 

				// Note that the set of properties fetched here must be aligned
				// with those set above.
				switch (property.tag()) {
				case PR_FID:
					data->fid = QString::number(property.value().toULongLong()); 
					break;
				case PR_MID:
					data->id = QString::number(property.value().toULongLong()); 
					break;
				case PR_CONVERSATION_TOPIC:
					data->title = property.value().toString(); 
					break;
				case PR_LAST_MODIFICATION_TIME:
					data->modified = property.value().toDateTime(); 
					break;
				default:
					qCritical() << "ignoring item property name:" << tagName(property.tag()) << property.value();
					break;
				}
			}

			// Add the entry to the output list!
			children.append(data);
			//TODO Just for debugging (in case the content list ist very long)
			//if (i >= 10) break;
		}
	}
	return true;
}

bool MapiFolder::open(mapi_id_t unused)
{
	Q_UNUSED(unused);
	mapi_id_t id = m_id;

	// Get the toplevel folder id if needed.
	if (id == 0) {
		if (MAPI_E_SUCCESS != GetDefaultFolder(m_connection->d(), &id, olFolderTopInformationStore)) {
			qCritical() << "cannot get default folder" << mapiError();
			return false;
		}
	}
	if (MAPI_E_SUCCESS != OpenFolder(m_connection->d(), id, &m_object)) {
		qCritical() << "cannot open folder" << id << mapiError();
		return false;
	}
	return true;
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
		qCritical() << "cannot create profile:" << mapiError();
		return false;
	}

	// TODO get workstation as parameter (was is it needed for anyway?)
	char hostname[256] = {};
	gethostname(&hostname[0], sizeof(hostname) - 1);
	hostname[sizeof(hostname) - 1] = 0;
	QString workstation = QString::fromLatin1(hostname);

	if (!addAttribute(profile8, "binding", server)) {
		qCritical() << "cannot add binding:" << server << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "workstation", workstation)) {
		qCritical() << "cannot add workstation:" << workstation << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "domain", domain)) {
		qCritical() << "cannot add domain:" << domain << mapiError();
		return false;
	}
// What is seal for? Seams to have something to do with Exchange 2010
// 	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "seal", (seal == true) ? "true" : "false");

// TODO Get langage from parameter if needed
// 	const char* locale = (const char *) (language) ? mapi_get_locale_from_language(language) : mapi_get_system_locale();
	const char *locale = mapi_get_system_locale();
	if (!locale) {
		qCritical() << "cannot find system locale:" << mapiError();
		return false;
	}

	uint32_t cpid = mapi_get_cpid_from_locale(locale);
	uint32_t lcid = mapi_get_lcid_from_locale(locale);
	if (!cpid || !lcid) {
		qCritical() << "Invalid Locale supplied or unknown system locale" << locale << ", deleting profile..." << mapiError();
		if (!remove(profile)) {
			return false;
		}
		return false;
	}

	if (!addAttribute(profile8, "codepage", QString::number(cpid))) {
		qCritical() << "cannot add codepage:" << cpid << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "language", QString::number(lcid))) {
		qCritical() << "cannot add codepage:" << cpid << mapiError();
		return false;
	}
	if (!addAttribute(profile8, "method", QString::number(lcid))) {
		qCritical() << "cannot add codepage:" << cpid << mapiError();
		return false;
	}


	struct mapi_session *session = NULL;
	if (MAPI_E_SUCCESS != MapiLogonProvider(m_context, &session, profile8, password.toUtf8(), PROVIDER_ID_NSPI)) {
		qCritical() << "cannot get logon provider, deleting profile..." << mapiError();
		if (!remove(profile)) {
			return false;
		}
		return false;
	}

	int retval = ProcessNetworkProfile(session, username.toUtf8().constData(), profileSelectCallback, NULL);
	if (retval != MAPI_E_SUCCESS && retval != 0x1) {
		qCritical() << "cannot process network profile, deleting profile..." << mapiError();
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

QString MapiProfiles::defaultGet()
{
	if (!init()) {
		return QString();
	}

	char *profname;
	if (MAPI_E_SUCCESS != GetDefaultProfile(m_context, &profname)) {
		qCritical() << "cannot get default profile:" << mapiError();
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
		qCritical() << "cannot set default profile:" << profile << mapiError();
		return false;
	}
	return true;
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
			qCritical() << "cannot make profile path:" << profilePath;
			return false;
		}
	}
	if (!QFile::exists(profileFile)) {
		if (MAPI_E_SUCCESS != CreateProfileStore(profileFile.toUtf8(), mapi_profile_get_ldif_path())) {
			qCritical() << "cannot create profile store:" << profileFile << mapiError();
			return false;
		}
	}
	if (MAPI_E_SUCCESS != MAPIInitialize(&m_context, profileFile.toLatin1())) {
		qCritical() << "cannot init profile store:" << profileFile << mapiError();
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
		qCritical() << "cannot get profile table" << mapiError();
		return QStringList();
	}

	// qDebug() << "Profiles in the database:" << proftable.cRows;
	QStringList profiles;
	for (unsigned count = 0; count < proftable.cRows; count++) {
		const char *name = proftable.aRow[count].lpProps[0].value.lpszA;
		uint32_t dflt = proftable.aRow[count].lpProps[1].value.l;

		profiles.append(QString::fromLocal8Bit(name));
		if (dflt) {
			qCritical() << "default profile:" << name;
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
		qCritical() << "cannot delete profile:" << profile << mapiError();
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
			qCritical() << "no default profile";
			return false;
		}
	}

	// Log on
	if (MAPI_E_SUCCESS != MapiLogonEx(m_context, &m_session, profile.toUtf8(), NULL)) {
		qCritical() << "cannot logon using profile" << profile << mapiError();
		return false;
	}
	if (MAPI_E_SUCCESS != OpenMsgStore(m_session, &m_store)) {
		qCritical() << "cannot open message store" << mapiError();
		return false;
	}
	return true;
}

bool MapiConnector2::resolveNames(const char *names[], SPropTagArray *tags,
				  SRowSet **results, PropertyTagArray_r **statuses)
{
	if (MAPI_E_SUCCESS != ResolveNames(m_session, names, tags, results, statuses, 0x0)) {
		qCritical() << "cannot resolve names" << mapiError();
		return false;
	}
	return true;
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

MapiMessage::MapiMessage(MapiConnector2 *connection, const char *tallocName, mapi_id_t id) :
	MapiObject(connection, tallocName, id)
{
	m_recipients.cRows = 0;
}

bool MapiMessage::open(mapi_id_t folderId)
{
	if (MAPI_E_SUCCESS != OpenMessage(m_connection->d(), folderId, m_id, &m_object, 0x0)) {
		qCritical() << "cannot open message:" << m_id << "in folder:" <<
			folderId << ", error:" << mapiError();
		return false;
	}
	return true;
}

unsigned MapiMessage::recipientCount() const
{
	return m_recipients.cRows;
}

bool MapiMessage::recipientsPull()
{
	SPropTagArray propertyTagArray;

	if (MAPI_E_SUCCESS != GetRecipientTable(&m_object, &m_recipients, &propertyTagArray)) {
		qCritical() << "cannot GetRecipientTable:" << mapiError();
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
		case PR_7BIT_DISPLAY_NAME:
		case PR_DISPLAY_NAME:
		case PR_RECIPIENT_DISPLAY_NAME:
			// We'd prefer a Unicode answer if there is one!
			if (result.name.isEmpty()) {
				result.name = property.value().toString(); 
			}
			break;
		case PR_7BIT_DISPLAY_NAME_UNICODE:
		case PR_DISPLAY_NAME_UNICODE:
		case PR_RECIPIENT_DISPLAY_NAME_UNICODE:
			result.name = property.value().toString(); 
			break;
		case PR_SMTP_ADDRESS:
			// We'd prefer a Unicode answer if there is one!
			if (result.email.isEmpty()) {
				result.email = property.value().toString(); 
			}
			break;
		case PR_SMTP_ADDRESS_UNICODE:
		case 0x6001001f:
			result.email = property.value().toString(); 
			break;
		case PR_RECIPIENT_TRACKSTATUS:
			result.trackStatus = property.value().toInt();
			break;
		case PR_RECIPIENT_FLAGS:
			result.flags = property.value().toInt();
			break;
		case PR_RECIPIENT_TYPE:
			result.type = property.value().toInt();
			break;
		case PR_RECIPIENT_ORDER:
			result.order = property.value().toInt();
			break;
		default:
			//qCritical() << "ignoring recipient property name:" << tagName(property.tag()) << property.value();
			break;
		}
	}
	return result;
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

bool MapiObject::propertyWrite(int tag, void *data, bool idempotent)
{
	// If the assignment is idempotent, if an instance of the 
	// property exists, it will be overwritten.
	if (idempotent) {
		for (unsigned i = 0; i < m_propertyCount; i++) {
			if (m_properties[i].ulPropTag == tag) {
				bool ok = set_SPropValue_proptag(&m_properties[i], (MAPITAGS)tag, data);
				if (!ok) {
					qCritical() << "cannot overwrite tag:" << tagName(tag) << "value:" << data;
				}
				return ok;
			}
		}
	}

	// Add a new entry to the array.
	m_properties = add_SPropValue(ctx(), m_properties, &m_propertyCount, (MAPITAGS)tag, data);
	if (!m_properties) {
		qCritical() << "cannot write tag:" << tagName(tag) << "value:" << data;
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
		qCritical() << "cannot talloc:" << data;
		return false;
	}

	return propertyWrite(tag, copy, idempotent);
}

bool MapiObject::propertyWrite(int tag, QDateTime &data, bool idempotent)
{
	FILETIME *copy = talloc(ctx(), FILETIME);

	if (!copy) {
		qCritical() << "cannot talloc:" << data;
		return false;
	}

	// As per http://support.citrix.com/article/CTX109645.
	time_t unixTime = data.toTime_t();
	NTTIME ntTime = (unixTime + 11644473600L) * 10000000;
	copy->dwHighDateTime = ntTime >> 32;
	copy->dwLowDateTime = ntTime;
	return propertyWrite(tag, copy, idempotent);
}

bool MapiObject::propertiesPush()
{
	if (MAPI_E_SUCCESS != SetProps(&m_object, MAPI_PROPS_SKIP_NAMEDID_CHECK, m_properties, m_propertyCount)) {
		qCritical() << "cannot push:" << m_propertyCount << "properties:" << mapiError();
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
				qCritical() << "cannot allocate tags:" << mapiError();
				return false;
			}
		} else {
			if (MAPI_E_SUCCESS != SPropTagArray_add(ctx(), tagArray, (MAPITAGS)tag)) {
				qCritical() << "cannot extend tags:" << mapiError();
				MAPIFreeBuffer(tagArray);
				return false;
			}
		}
	}
	if (MAPI_E_SUCCESS != GetProps(&m_object, MAPI_UNICODE, tagArray, &m_properties, &m_propertyCount)) {
		qCritical() << "cannot pull properties:" << mapiError();
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
		qCritical() << "cannot pull all properties:" << mapiError();
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
		qCritical() << "cannot copy properties:" << mapiError();
	}
	return true;
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

QString MapiObject::tagAt(unsigned i) const
{
	if (!m_propertyCount || ((m_propertyCount - 1) < i)) {
		return QString();
	}
	return tagName(m_properties[i].ulPropTag);
}

QString MapiObject::propertyString(unsigned i) const
{
	if (!m_propertyCount || ((m_propertyCount - 1) < i)) {
		return QString();
	}
	return MapiProperty(m_properties[i]).toString();
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

MapiProperty::MapiProperty(SPropValue &property) :
	m_property(property)
{
}

/**
 * Get the value of the property in a nice typesafe wrapper.
 */
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

int MapiProperty::tag() const
{
	return m_property.ulPropTag;
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
		qDebug() << "unsupported frequency:"<<pattern->RecurFrequency;
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
			qDebug() << "unsupported endtype:"<<pattern->EndType;
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
