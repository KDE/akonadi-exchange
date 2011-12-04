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

MapiConnector2::MapiConnector2()
:m_context(NULL), m_session(NULL)
{
	if (MAPIInitialize(&m_context, getMapiProfileDirectory().toLatin1()) != MAPI_E_SUCCESS) {
		return;
	}
}

MapiConnector2::~MapiConnector2()
{
	if (m_session!=NULL) {
		Logoff(&m_store);
	}

	MAPIUninitialize(m_context);
}

QString MapiConnector2::getMapiProfileDirectory()
{
	QString profilePath(QDir::home().path());
	profilePath.append(QLatin1String("/.openchange/"));
	
	QString profileFile(profilePath);
	profileFile.append(QString::fromLatin1("profiles.ldb"));

	// check if the store exists
	QDir path(profilePath);
	if (!path.exists()) {
		bool ok = path.mkpath(profilePath);
		if (!ok) {
			return QString();
		}
	}
	if (!QFile::exists(profileFile)) {
		QLatin1String ldif(mapi_profile_get_ldif_path());
		MAPISTATUS retval = CreateProfileStore(profileFile.toUtf8().constData(), ldif.latin1());
        if (retval != MAPI_E_SUCCESS) {
			return QString();
        }
	}

	return profileFile;
}

bool MapiConnector2::removeProfile(QString profile)
{
	enum MAPISTATUS retval;
	if ((retval = DeleteProfile(m_context, profile.toUtf8().constData()) ) != MAPI_E_SUCCESS) {
		return false;
	}
	return true;
}

int profileSelectCallback(struct SRowSet *rowset, const void* /*private_var*/)
{
  qDebug() << "Found more than 1 matching users -> cancel";

  //  TODO Some sort of handling would be needed here
  return rowset->cRows;
}

bool MapiConnector2::createProfile(QString profile, QString username, QString password, QString domain, QString server)
{
	enum MAPISTATUS retval;

	retval = CreateProfile(m_context, profile.toUtf8().constData(), username.toUtf8().constData(), password.toUtf8().constData(), 0);
	if (retval != MAPI_E_SUCCESS) {
		return false;
	}

	// TODO get workstation as parameter (was is it needed for anyway?)
	char hostname[256] = {};
	const char* workstation = NULL;
	if (!workstation) {
		gethostname(hostname, sizeof(hostname) - 1);
		hostname[sizeof(hostname) - 1] = 0;
		workstation = hostname;
	}

	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "binding", server.toUtf8().constData());
	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "workstation", workstation);
	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "domain", domain.toUtf8().constData());
// What is seal for? Seams to have something to do with Exchange 2010
	// 	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "seal", (seal == true) ? "true" : "false");

// TODO Get langage from parameter if needed
// 	const char* locale = (const char *) (language) ? mapi_get_locale_from_language(language) : mapi_get_system_locale();
	const char* locale = mapi_get_system_locale();
	if (locale == NULL) {
		qDebug() << "Unable to find system locale";
		return false;
	}

	uint32_t cpid = mapi_get_cpid_from_locale(locale);
	uint32_t lcid = mapi_get_lcid_from_locale(locale);

	if (!cpid || !lcid) {
		qDebug() << "Invalid Locale supplied or unknown system locale" << locale << "; Deleting profile...";
		if ((retval = DeleteProfile(m_context, profile.toUtf8().constData())) != MAPI_E_SUCCESS) {
			return false;
		}
		return false;
	}

	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "codepage", QString::number(cpid).toUtf8().constData());
	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "language", QString::number(lcid).toUtf8().constData());
	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "method", QString::number(lcid).toUtf8().constData());


	struct mapi_session *session = NULL;
	retval = MapiLogonProvider(m_context, &session, profile.toUtf8().constData(), password.toUtf8().constData(), PROVIDER_ID_NSPI);
	if (retval != MAPI_E_SUCCESS) {
		if ((retval = DeleteProfile(m_context, profile.toUtf8().constData())) != MAPI_E_SUCCESS) {
			return false;
		}
		return false;
	}

	retval = ProcessNetworkProfile(session, username.toUtf8().constData(), profileSelectCallback, NULL);

	if (retval != MAPI_E_SUCCESS && retval != 0x1) {
		if ((retval = DeleteProfile(m_context, profile.toUtf8().constData())) != MAPI_E_SUCCESS) {
			return false;
		}
		return false;
	}

	return true;
}

bool MapiConnector2::setDefaultProfile(QString profile)
{
	enum MAPISTATUS retval;
	if ( (retval = SetDefaultProfile(m_context, profile.toUtf8().constData())) != MAPI_E_SUCCESS) {
		return false;
	}
	return true;
}

QString MapiConnector2::getDefaultProfile()
{
	char * profname;
	MAPISTATUS retval = GetDefaultProfile(m_context, &profname);
	if (retval != MAPI_E_SUCCESS) {
		return QString();
	}
	return QString::fromLocal8Bit(profname);
}

bool MapiConnector2::login(QString profilename)
{
	if (m_session != NULL) {
		// already logged in...
		return true;
	}

	enum MAPISTATUS retval;

	if (profilename.isEmpty()) {
		// use the default profile if none was specified by the caller
		profilename = getDefaultProfile();
		if (profilename.isEmpty()) {
			// there seams to be no default profile
			return false;
		}
	}

	// Log on
	retval = MapiLogonEx(m_context, &m_session, profilename.toUtf8().constData(), NULL);
	if (retval != MAPI_E_SUCCESS) {
		return false;
	}

	mapi_object_init(&m_store);
	retval = OpenMsgStore(m_session, &m_store);
	if (retval != MAPI_E_SUCCESS) {
		return false;
	}

	return true;
}

QStringList MapiConnector2::listProfiles()
{
	enum MAPISTATUS retval;
	struct SRowSet  proftable;
	uint32_t        count = 0;

	memset(&proftable, 0, sizeof (struct SRowSet));
	if ((retval = GetProfileTable(m_context, &proftable)) != MAPI_E_SUCCESS) {
		return QStringList();
	}

	// qDebug() << "Profiles in the database:" << proftable.cRows;

	QStringList profileList;
	for (count = 0; count != proftable.cRows; count++) {
		const char* name = proftable.aRow[count].lpProps[0].value.lpszA;
		uint32_t dflt = proftable.aRow[count].lpProps[1].value.l;

		profileList.append(QString::fromLocal8Bit(name));
		if (dflt) {
			// TODO
		}
	}

	return profileList;
}

bool MapiConnector2::fetchFolderList(QList< FolderData >& list, mapi_id_t parentFolderID, const QString filter)
{
	mapi_id_t topFolderID;
	if (parentFolderID == 0) {
		// get the toplevel folder
		if (GetDefaultFolder(&m_store, &topFolderID, olFolderTopInformationStore) != MAPI_E_SUCCESS) {
			return false;
		}
	} else {
		// or use the supplied parent folder
		topFolderID = parentFolderID;
	}
	mapi_object_t topFolder = openFolder(topFolderID);

	// get all sub folders
	mapi_object_t hierarchy_table;

	mapi_object_init(&hierarchy_table);
	if (GetHierarchyTable(&topFolder, &hierarchy_table, 0, NULL) != MAPI_E_SUCCESS) {
		mapi_object_release(&hierarchy_table);
		return false;
	}

	SPropTagArray* property_tag_array = set_SPropTagArray(m_context, 0x3, PR_FID, PR_DISPLAY_NAME, PR_CONTAINER_CLASS);
	if (SetColumns(&hierarchy_table, property_tag_array)) {
		MAPIFreeBuffer(property_tag_array);
		mapi_object_release(&hierarchy_table);
		return false;
	}

	MAPIFreeBuffer(property_tag_array);

	/* Get current cursor position */
	uint32_t Denominator;
	if (QueryPosition(&hierarchy_table, NULL, &Denominator) != MAPI_E_SUCCESS) {
		mapi_object_release(&hierarchy_table);
		return false;
	}

	SRowSet rowset;
	while( (QueryRows(&hierarchy_table, Denominator, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS) && rowset.cRows) {
		for (unsigned int i = 0; i < rowset.cRows; ++i) {

			mapi_id_t* fid = (mapi_id_t *)find_SPropValue_data(&(rowset.aRow[i]), PR_FID);
			const char* name = (const char *)find_SPropValue_data(&(rowset.aRow[i]), PR_DISPLAY_NAME);
			const char* folderClass = (const char *)find_SPropValue_data(&(rowset.aRow[i]), PR_CONTAINER_CLASS);

			if (folderClass!=NULL && !filter.isEmpty()) {
				if (!QString::fromAscii(folderClass).startsWith(filter)) {
					qDebug() << "Folder"<<name<<"does not match filter"<<filter;
					continue;
				}
			}

			FolderData data;
			data.id = QString::number(*fid);
			data.name = QString::fromAscii(name);
			list.append(data);
		}
	}

	mapi_object_release(&hierarchy_table);

	return true;
}

bool MapiConnector2::fetchFolderContent(mapi_id_t folderID, QList<CalendarDataShort>& list)
{
	mapi_object_t topFolder = openFolder(folderID);

	// Retrieve folder's content table
	mapi_object_t obj_table;
	mapi_object_init(&obj_table);
	if (GetContentsTable(&topFolder, &obj_table, 0x0, NULL) != MAPI_E_SUCCESS) {
		mapi_object_release(&obj_table);
		return false;
	}

	// Create the MAPI table view
	struct SPropTagArray *SPropTagArray =  set_SPropTagArray(m_context, 0x4, PR_FID, PR_MID, PR_CONVERSATION_TOPIC, PR_LAST_MODIFICATION_TIME);
	if (SetColumns(&obj_table, SPropTagArray) != MAPI_E_SUCCESS) {
		MAPIFreeBuffer(SPropTagArray);
		mapi_object_release(&obj_table);
		return false;
	}
	MAPIFreeBuffer(SPropTagArray);

	// Get current cursor position
	uint32_t Denominator;
	if (QueryPosition(&obj_table, NULL, &Denominator) != MAPI_E_SUCCESS) {
		mapi_object_release(&obj_table);
		return false;
	}

	// Iterate through rows
	SRowSet rowset;
	while (QueryRows(&obj_table, Denominator, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS && rowset.cRows) {
		for (uint32_t i = 0; i < rowset.cRows; i++) {
			mapi_id_t* fid = (mapi_id_t *)find_SPropValue_data(&(rowset.aRow[i]), PR_FID);
			mapi_id_t* mid = (mapi_id_t *)find_SPropValue_data(&(rowset.aRow[i]), PR_MID);
			const char* topic = (const char *)find_SPropValue_data(&(rowset.aRow[i]), PR_CONVERSATION_TOPIC);
			const FILETIME* modTime = (const FILETIME*)find_SPropValue_data(&(rowset.aRow[i]), PR_LAST_MODIFICATION_TIME);

			CalendarDataShort data;
			data.fid = QString::number(*fid);
			data.id = QString::number(*mid);
			data.title = QString::fromAscii(topic);
			data.modified = convertSysTime(*modTime);
			list.append(data);

			//TODO Just for debugging (in case the content list ist very long)
			//if (i >= 10) break;
		}
	}

	mapi_object_release(&obj_table);
	return true;
}

bool MapiConnector2::fetchCalendarData(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data)
{
	TallocContext ctx("MapiConnector2::fetchCalendarData");
	MapiObject folder(ctx, folderID);
	MapiAppointment message(ctx, messageID);

	if (!message.open(&m_store, folderID)) {
		return false;
	}

	if (!message.propertiesPull()) {
		return false;
	}

	data.fid = QString::number(folderID);
	data.id = QString::number(messageID);
	data.title = message.property(PidTagConversationTopic).toString();
	data.text = message.property(PidTagBody).toString();
	data.modified = message.property(PidTagLastModificationTime).toDateTime();
	data.created = message.property(PidTagCreationTime).toDateTime();
	data.begin = message.property(PidTagStartDate).toDateTime();
	data.end = message.property(PidTagEndDate).toDateTime();
	data.sender = message.property(PidTagSentRepresentingName).toString();
	data.location = message.property(PidLidLocation).toString();
//	const char* recurrencePattern = (const char* )find_mapi_SPropValue_data(&properties_array, PidLidRecurrencePattern);
//	const SBinary_short *binData = (const SBinary_short*)find_mapi_SPropValue_data(&properties_array, PidLidAppointmentRecur);
	RecurrencePattern *pattern = message.recurrance();
	if (pattern) {
		data.recurrency.setData(pattern);
	}
	data.reminderActive = message.property(PidLidReminderSet).toInt();
	if (data.reminderActive) {
		data.reminderTime = message.property(PidLidReminderSignalTime).toDateTime();
		data.reminderDelta = message.property(PidLidReminderDelta).toInt();
	}
	QString toAttendeesString = message.property(PidTagDisplayTo).toString();
	getAttendees(*message.d(), toAttendeesString, data);
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

TALLOC_CTX *TallocContext::d()
{
	return m_ctx;
}

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

class MapiProperty: private SPropValue
{
public:
	MapiProperty(SPropValue &property) :
		m_property(property)
	{
	}

	/**
	 * Get the value of the property in a nice typesafe wrapper.
	 */
	QVariant value() const
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
	QString toString() const
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

private:
	SPropValue &m_property;
};

RecurrencePattern *MapiAppointment::recurrance()
{
	unsigned recurrenceType = propertyFind(PidLidRecurrenceType);

	// Is there a recurrance type set?
	if (recurrenceType == UINT_MAX) {
		return 0;
	}
	recurrenceType = propertyAt(recurrenceType).toInt();
	if (recurrenceType == 0x0) {
		return 0;
	}

	// Can we find the recurrance data? 
	unsigned recurrenceData = propertyFind(PidLidAppointmentRecur);
	if (recurrenceData == UINT_MAX) {
		// TODO This should not happen. PidLidRecurrenceType says this is a recurring event, so why is there no PidLidAppointmentRecur???
		qDebug() << "missing recurrencePattern in message" << m_id;
		return 0;
	}
	RecurrencePattern *pattern = get_RecurrencePattern(m_ctx.d(), &m_properties[recurrenceData].value.bin);
	debugRecurrencyPattern(pattern);
	return pattern;
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

MapiObject::MapiObject(TallocContext &ctx, mapi_id_t id) :
	m_ctx(ctx),
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

mapi_object_t *MapiObject::d()
{
	return &m_object;
}

mapi_id_t MapiObject::id()
{
	return m_id;
}

bool MapiObject::open(mapi_object_t *store, mapi_id_t folderId)
{
	if (MAPI_E_SUCCESS != OpenMessage(store, folderId, m_id, &m_object, 0x0)) {
		qCritical() << "cannot open message:" << m_id << "in folder:" << folderId
		<< ", error:" << mapiError();
		return false;
	}
	return true;
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
	m_properties = add_SPropValue(m_ctx.d(), m_properties, &m_propertyCount, (MAPITAGS)tag, data);
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
	char *copy = talloc_strdup(m_ctx.d(), data.toUtf8().data());

	if (!copy) {
		qCritical() << "cannot talloc:" << data;
		return false;
	}

	return propertyWrite(tag, copy, idempotent);
}

bool MapiObject::propertyWrite(int tag, QDateTime &data, bool idempotent)
{
	FILETIME *copy = talloc(m_ctx.d(), FILETIME);

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
			tagArray = set_SPropTagArray(m_ctx.d(), 1, (MAPITAGS)tag);
			if (!tagArray) {
				qCritical() << "cannot allocate tags:" << mapiError();
				return false;
			}
		} else {
			if (MAPI_E_SUCCESS != SPropTagArray_add(m_ctx.d(), tagArray, (MAPITAGS)tag)) {
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
	m_properties = talloc_array(m_ctx.d(), SPropValue, mapiProperties.cValues);
	if (m_properties) {
		for (m_propertyCount = 0; m_propertyCount < mapiProperties.cValues; m_propertyCount++) {
			cast_SPropValue(m_ctx.d(), &mapiProperties.lpProps[m_propertyCount], 
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

QString MapiObject::tagAt(unsigned i)
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

QString MapiObject::tagName(int tag)
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

MapiAppointment::MapiAppointment(TallocContext &ctx, mapi_id_t id) :
	MapiObject(ctx, id)
{
}

bool MapiAppointment::open(mapi_object_t *store, mapi_id_t folderId)
{
	if (!MapiObject::open(store, folderId)) {
		return false;
	}

	// Sanity check the message class.
	QVector<int> readTags;
	readTags.append(PidTagMessageClass);
	if (!propertiesPull(readTags)) {
		return false;
	}
	QString messageClass = property(PidTagMessageClass).toString();
	if (QString::fromLatin1("IPM.Appointment").compare(messageClass)) {
		// this one is not an appointment
		return false;
	}
	return true;
}

bool MapiConnector2::calendarDataUpdate(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data)
{
	TallocContext ctx("MapiConnector2::calendarDataUpdate");
	MapiObject folder(ctx, folderID);
	MapiAppointment message(ctx, messageID);

	if (!message.open(&m_store, folderID)) {
		return false;
	}

        if (!message.propertiesPull()) {
		return false;
	}
	for (unsigned i = 0; i < message.propertyCount(); i++) {
		qCritical() << "pulled:" << message.tagAt(i) << "value:" << message.propertyString(i); 
	}

	// Overwrite all the fields we know about.
	if (!message.propertyWrite(PidTagConversationTopic, data.title)) {
		return false;
	}
	if (!message.propertyWrite(PidTagBody, data.text)) {
		return false;
	}
	if (!message.propertyWrite(PidTagLastModificationTime, data.modified)) {
		return false;
	}
	if (!message.propertyWrite(PidTagCreationTime, data.created)) {
		return false;
	}
	if (!message.propertyWrite(PidTagStartDate, data.begin)) {
		return false;
	}
	if (!message.propertyWrite(PidTagEndDate, data.end)) {
		return false;
	}
	if (!message.propertyWrite(PidTagSentRepresentingName, data.sender)) {
		return false;
	}
	if (!message.propertyWrite(PidLidLocation, data.location)) {
		return false;
	}
	if (!message.propertyWrite(PidLidReminderSet, data.reminderActive)) {
		return false;
	}
	if (data.reminderActive) {
		if (!message.propertyWrite(PidLidReminderSignalTime, data.reminderTime)) {
			return false;
		}
		if (!message.propertyWrite(PidLidReminderDelta, data.reminderDelta)) {
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
	if (!message.propertiesPush()) {
		return false;
	}
#if 0
	qCritical() << "************  OpenFolder";
	if (MAPI_E_SUCCESS != OpenFolder(&m_store, folder.id(), folder.d())) {
		qCritical() << "cannot open folder" << folderID
			<< ", error:" << mapiError();
		return false;
        }
	qCritical() << "************  SaveChangesMessage";
	if (MAPI_E_SUCCESS != SaveChangesMessage(folder.d(), message.d(), KeepOpenReadWrite)) {
		qCritical() << "cannot save message" << messageID << "in folder" << folderID
			<< ", error:" << mapiError();
		return false;
	}
#endif
	qCritical() << "************  SubmitMessage";
	if (MAPI_E_SUCCESS != SubmitMessage(message.d())) {
		qCritical() << "cannot submit message" << messageID << "in folder" << folderID
			<< ", error:" << mapiError();
		return false;
	}
	struct mapi_SPropValue_array replyProperties;
	qCritical() << "************  TransportSend";
	if (MAPI_E_SUCCESS != TransportSend(message.d(), &replyProperties)) {
		qDebug() << "cannot send message" << messageID << "in folder" << folderID
			<< ", error:" << mapiError();
		return false;
	}
	return true;
}

mapi_object_t MapiConnector2::openFolder(mapi_id_t folderID)
{
	enum MAPISTATUS retval;

	mapi_object_t obj_folder;
	mapi_object_init(&obj_folder);
	retval = OpenFolder(&m_store, folderID, &obj_folder);
	if (retval != MAPI_E_SUCCESS) {
		mapi_object_release(&obj_folder);
		return obj_folder;
	}

	return obj_folder;
}

QString MapiConnector2::mapiValueToQString(mapi_SPropValue *lpProps)
{
	QString dataStr;
	switch(lpProps->ulPropTag & 0xFFFF) {
		case PT_SHORT:
			dataStr = QString::number(lpProps->value.i); break;
		case PT_BOOLEAN:
			dataStr = QString::number(lpProps->value.b); break;
		case PT_I8:
			dataStr = QString::number(lpProps->value.d); break;
		case PT_STRING8:
			dataStr = QString::fromLocal8Bit(lpProps->value.lpszA).append(QString::fromAscii(" (PT_STRING8)")); break;
		case PT_UNICODE:
			dataStr = QString::fromUtf8(lpProps->value.lpszW).append(QString::fromAscii(" (PT_UNICODE)")); break;
		case PT_SYSTIME:
			dataStr = convertSysTime(lpProps->value.ft).toString(); break;
		case PT_ERROR:
			dataStr = QString::number(lpProps->value.err); break;
		case PT_LONG:
			dataStr = QString::number(lpProps->value.l); break;
		case PT_DOUBLE:
			dataStr = QString::number(lpProps->value.dbl); break;
		case PT_CLSID:
			dataStr = QString::fromAscii("UNSUPPORTED(PT_CLSID)"); break;
// 					return (const void *)lpProps->value.lpguid;
		case PT_BINARY:
			dataStr = QString::fromAscii("UNSUPPORTED(PT_BINARY)"); break;
// 					return (const void *)&lpProps->value.bin;
		case PT_OBJECT:
			dataStr = QString::fromAscii("UNSUPPORTED(PT_OBJECT)"); break;
// 					return (const void *)&lpProps->value.object;
		default:
			dataStr = QString::fromAscii("UNKNOWN TYPE"); break;
	}
	return dataStr;
}


bool MapiConnector2::getAttendees(mapi_object_t& obj_message, const QString& toAttendeesStr, CalendarData& data)
{
	QStringList attendeesNames = toAttendeesStr.split(QString::fromAscii("; "));

	SPropTagArray propertyTagArray;
	SRowSet rowSet;
	if ( GetRecipientTable( &obj_message, &rowSet, &propertyTagArray ) != MAPI_E_SUCCESS ) {
		qDebug() << "Error opening GetRecipientTable()";
		return false;
	}

	for (uint32_t i = 0; i < rowSet.cRows; ++i) {
		QString name, email;
		uint32_t *trackStatus = NULL, *flags = NULL, *type = NULL, *number=NULL;

// 		qDebug() << "Data for attendee number"<<i<<"valuecount:"<<rowSet.aRow[i].cValues;

		for (unsigned int j = 0; j < rowSet.aRow[i].cValues; ++j) {
			switch ( rowSet.aRow[i].lpProps[j].ulPropTag ) {
			case 0x5ff6001f: // Should be PR_DISPLAY_NAME
			case PR_7BIT_DISPLAY_NAME:
				if (name.isEmpty())
					name = QString::fromLocal8Bit( rowSet.aRow[i].lpProps[j].value.lpszA );
				break;
			case 0x39fe001f: // Should be PR_SMTP_ADDRESS
				email = QString::fromLocal8Bit( rowSet.aRow[i].lpProps[j].value.lpszA );
			break;
			case PR_RECIPIENT_TRACKSTATUS:
				trackStatus = &rowSet.aRow[i].lpProps[j].value.l;
				break;
			case PR_RECIPIENT_FLAGS:
				flags = &rowSet.aRow[i].lpProps[j].value.l;
				break;
			case PR_RECIPIENT_TYPE:
				type = &rowSet.aRow[i].lpProps[j].value.l;
				break;
			case PR_RECIPIENT_ORDER:
				number = &rowSet.aRow[i].lpProps[j].value.l;
				break;
			default:
				break;
			}

// TODO for debugging
// 			QString tagStr;
// 			tagStr.sprintf("0x%X", rowSet.aRow[i].lpProps[j].ulPropTag);
// 			qDebug() << tagStr << mapiValueToQString(&rowSet.aRow[i].lpProps[j]);
		}

		uint32_t idx = (number)?*number:0;

		Attendee attendee;
		attendee.name = (!name.isEmpty())?name:QString();
		attendee.email = (!email.isEmpty())?email:name;
		if (flags && (*flags & 0x0000002)) {
			attendee.isOranizer=true;
		} else {
			attendee.isOranizer=false;
		}
		attendee.status = (trackStatus)?*trackStatus:0;
		attendee.type = (type)?*type:0;
		attendee.idx = idx;
		data.anttendees << attendee;
	}

	// check if we need the lookup the real e-mail adresses
	QStringList namesNeedResolve;
	foreach (const Attendee& att, data.anttendees) {
		if (!att.email.isEmpty() && !att.email.contains(QString::fromAscii("@"))) {
			namesNeedResolve << att.email;
		}
	}
	QMap<QString, RecipientData> recipientData;
	if (namesNeedResolve.size() > 0) {
		resolveNames(namesNeedResolve, recipientData);
	}

	for (int i=0; i<data.anttendees.size(); ++i) {
		// update name and email with the values gained from resolve name
		QMap<QString, RecipientData>::const_iterator it;
		it =recipientData.find(data.anttendees[i].email); 
		if (it!=recipientData.end()) {
			data.anttendees[i].name = it->name;
			data.anttendees[i].email = it->email;
			continue; // update done -> move on to next one
		}

		// if the e-mail adress is still not present take if from the to-attendees list
		if (data.anttendees[i].email.isEmpty()) {
			if (data.anttendees[i].type == MAPI_TO) {
				if (data.anttendees[i].idx != 0) {
					data.anttendees[i].email = attendeesNames.at(data.anttendees[i].idx-1);
				} else {
					data.anttendees[i].email = attendeesNames.at(i);
				}
			}
		}
	}

	foreach (const Attendee& attendee, data.anttendees) {
		qDebug()<<"Attendee"<<attendee.name<<attendee.email<<attendee.status<<attendee.type;
	}

	return true;
}

void MapiConnector2::resolveNames(const QStringList& names, QMap<QString, RecipientData>& outputMap) 
{
// 	qDebug()<<"resolving names"<<names;

	const char* nameArray[names.size()+1];
	int i=0;
	foreach (const QString name, names) {
		nameArray[i] = strdup( name.toStdString().c_str() );
		++i;
	}
	nameArray[i] = NULL;

	struct PropertyTagArray_r* flaglist = NULL;
	struct SPropTagArray *SPropTagArray = NULL;
	struct SRowSet *SRowSet = NULL;

	TALLOC_CTX *mem_ctx;
	mem_ctx = talloc_named(NULL, 0, "MapiConnector2::resolveNames");

	// TODO why no UNICODE here?
	SPropTagArray = set_SPropTagArray (mem_ctx, 0x2,
					PR_DISPLAY_NAME,
					PR_SMTP_ADDRESS);

 	if (ResolveNames (this->m_session, nameArray, SPropTagArray, &SRowSet, &flaglist, 0x0) == MAPI_E_SUCCESS) {
		if (SRowSet) {
			for (uint32_t i = 0; i < SRowSet->cRows; ++i) {
				QString name, email;
				for (unsigned int j = 0; j < SRowSet->aRow[i].cValues; ++j) {
					switch ( SRowSet->aRow[i].lpProps[j].ulPropTag ) {
						case PR_DISPLAY_NAME:
							if (name.isEmpty())
								name = QString::fromLocal8Bit( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						case PR_SMTP_ADDRESS:
							email = QString::fromLocal8Bit( SRowSet->aRow[i].lpProps[j].value.lpszA );
							break;
						default:
							break;
					}
// 					TODO just for debugging
// 					QString tagStr;
// 					tagStr.sprintf("0x%X", SRowSet->aRow[i].lpProps[j].ulPropTag);
// 					qDebug() << tagStr << mapiValueToQString(&SRowSet->aRow[i].lpProps[j]);
				}
				RecipientData data;
				data.name = name;
				data.email = email;
				outputMap.insert(QString::fromAscii(nameArray[i]), data);
			}
		}
	}
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
