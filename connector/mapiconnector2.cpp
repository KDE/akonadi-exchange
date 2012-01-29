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

#include "mapiconnector2.h"

#include <QDebug>
#include <QStringList>
#include <QDir>
#include <QMessageBox>
#include <QRegExp>
#include <QVariant>

#include <kpimutils/email.h>

/**
 * Display ids in the same format we use when stored in Akonadi.
 */
#define ID_BASE 36

#define CASE_PREFER_A_OVER_B(a, b, lvalue, rvalue) \
case b: \
	if (lvalue.isEmpty()) { \
		lvalue = rvalue; \
	} \
	break; \
case a: \
	lvalue = rvalue; \

#define UNDOCUMENTED_PR_EMAIL_UNICODE 0x6001001f

/**
 * Set this to 1 to pull all the properties, e.g. to see what a server has
 * available.
 */
#ifndef DEBUG_MESSAGE_PROPERTIES
#define DEBUG_MESSAGE_PROPERTIES 0
#endif

#define STR(def) \
case def: return QString::fromLatin1(#def)

/**
 * Map all MAPI errors to strings.
 */
QString mapiError()
{
	int code = GetLastError();
	switch (code)
	{
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
	default:
		return QString::fromAscii("MAPI_E_0x%1").arg((unsigned)code, 0, 16);
	}
}

/**
 * Try to extract an email address from a string.
 */
extern QString mapiExtractEmail(const QString &source, const QByteArray &type, bool emptyDefault)
{
	QString email;
	
	if (!emptyDefault) {
		email = source;
	}
	if (type == "SMTP") {
		QString name;

		// First, we give the library routines a chance.
		if (!KPIMUtils::extractEmailAddressAndName(source, email, name)) {
			// Now for some custom action. Look for the last possible
			// starting delimiter, and work forward from there. Thus:
			//
			//       "blah (blah) <blah> <result>"
			//
			// should return "result".
			static QRegExp firstRE(QString::fromAscii("[(<]"));
			static QRegExp lastRE(QString::fromAscii("[)>]"));

			int first = source.lastIndexOf(firstRE);
			int last = source.indexOf(lastRE, first);

			if ((first > -1) && (last > first + 1)) {
				email = source.mid(first + 1, last - first - 1);
			}
		}
	} else if (type == "EX") {
		// Convert an "EX"change address to an account name, which 
		// should be the email alias.
		int lastCn = source.lastIndexOf(QString::fromAscii("/CN="), -1, Qt::CaseInsensitive);

		if (lastCn > -1) {
			email = source.mid(lastCn + 4);
		}
	}
	return email;
}

QString mapiExtractEmail(const class MapiProperty &source, const QByteArray &type, bool emptyDefault)
{
	return mapiExtractEmail(source.value().toString(), type, emptyDefault);
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

static QChar x500Prefix(QChar::fromAscii('/'));
static QChar domainSeparator(QChar::fromAscii('@'));

unsigned isGoodEmailAddress(QString &email)
{

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
	if (!email.contains(domainSeparator)) {
		return 2;
	} else {
		return 3;
	}
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

bool MapiConnector2::GALCount(unsigned *totalCount)
{
	if (MAPI_E_SUCCESS != GetGALTableCount(m_session, totalCount)) {
		error() << "cannot get GAL count" << mapiError();
		return false;
	}
	return true;
}

bool MapiConnector2::GALRead(unsigned requestedCount, SPropTagArray *tags, SRowSet **results, unsigned *percentagePosition)
{
	if (MAPI_E_SUCCESS != GetGALTable(m_session, tags, results, requestedCount, TABLE_CUR)) {
		error() << "cannot read GAL entries" << mapiError();
		return false;
	}

	// Return where we got to.
	if (percentagePosition) {
		struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;

		*percentagePosition = nspi->pStat->NumPos * 100 / nspi->pStat->TotalRecs;
	}
	return true;
}

bool MapiConnector2::GALRewind()
{
	struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;

	nspi->pStat->CurrentRec = 0;
	nspi->pStat->Delta = 0;
	nspi->pStat->NumPos = 0;
	nspi->pStat->TotalRecs = 0xffffffff;
	return true;
}

bool MapiConnector2::GALSeek(const QString &displayName, unsigned *percentagePosition, SPropTagArray *tags, SRowSet **results)
{
	struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;
	SPropValue key;
	SRowSet *dummy;

	key.ulPropTag = (MAPITAGS)PR_DISPLAY_NAME_UNICODE;
	key.dwAlignPad = 0;
	key.value.lpszW = string(displayName);
	if (MAPI_E_SUCCESS != nspi_SeekEntries(nspi, ctx(), SortTypeDisplayName, &key, tags, NULL, results ? results : &dummy)) {
		error() << "cannot seek to GAL entry" << displayName << mapiError();
		return false;
	}

	// Return where we got to.
	if (percentagePosition) {
		*percentagePosition = nspi->pStat->NumPos * 100 / nspi->pStat->TotalRecs;
	}
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
	name = QString::number(id, ID_BASE);
}

MapiFolder::~MapiFolder()
{
	mapi_object_release(&m_contents);
}

QDebug MapiFolder::debug() const
{
	static QString prefix = QString::fromAscii("MapiFolder: %1:");
	return MapiObject::debug(prefix.arg(m_id, 0, ID_BASE));
}

QDebug MapiFolder::error() const
{
	static QString prefix = QString::fromAscii("MapiFolder: %1:");
	return MapiObject::error(prefix.arg(m_id, 0, ID_BASE));
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
					//debug() << "ignoring folder property:" << tagName(property.tag()) << property.value();
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
					//debug() << "ignoring item property:" << tagName(property.tag()) << property.value();
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
	if (MAPI_E_SUCCESS != OpenFolder(m_connection->d(), m_id, &m_object)) {
		error() << "cannot open folder" << m_id << mapiError();
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
}

void MapiMessage::addUniqueRecipient(const char *source, MapiRecipient &candidate)
{
	debug() << "candidate address:" << source << candidate.toString();
	if (candidate.name.isEmpty() && candidate.email.isEmpty()) {
		// Discard garbage.
		return;
	}

	for (int i = 0; i < m_recipients.size(); i++) {
		MapiRecipient &entry = m_recipients[i];

		// A ReplyTo item only matches other ReplyTo items.
		if ((candidate.type() == MapiRecipient::ReplyTo) && (entry.type() != MapiRecipient::ReplyTo)) {
			continue;
		}

		// If we find a name match, fill in a missing email if we can.
		if (!candidate.name.isEmpty() && (0 == entry.name.compare(candidate.name, Qt::CaseInsensitive))) {
			if (isGoodEmailAddress(entry.email) < isGoodEmailAddress(candidate.email)) {
				entry.email = candidate.email;
				// Promote the type if needed to more specific
				// (numerically lower) type.
				if (entry.type() > candidate.type()) {
					entry.setType(candidate.type());
				}
				// Promote the object and display type to a non-default value.
				if (entry.displayType() == MapiRecipient::DtMailuser) {
					entry.setDisplayType(candidate.displayType());
				}
				if (entry.objectType() == MapiRecipient::OtMailuser) {
					entry.setObjectType(candidate.objectType());
				}
			}
			return;
		}

		// If we find an email match, fill in a missing name if we can.
		if (!candidate.email.isEmpty() && (0 == entry.email.compare(candidate.email, Qt::CaseInsensitive))) {
			if (entry.name.length() < candidate.name.length()) {
				entry.name = candidate.name;
				// Promote the type if needed to more specific
				// (numerically lower) type.
				if (entry.type() > candidate.type()) {
					entry.setType(candidate.type());
				}
				// Promote the object and displkay type to a non-default value.
				if (entry.displayType() == MapiRecipient::DtMailuser) {
					entry.setDisplayType(candidate.displayType());
				}
				if (entry.objectType() == MapiRecipient::OtMailuser) {
					entry.setObjectType(candidate.objectType());
				}
			}
			return;
		}
	}

	// Add the entry if it did not match.
	debug() << "add new address:" << source << candidate.toString();
	m_recipients.append(candidate);
}

QDebug MapiMessage::debug() const
{
	static QString prefix = QString::fromAscii("MapiMessage: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE));
}

QDebug MapiMessage::error() const
{
	static QString prefix = QString::fromAscii("MapiMessage: %1/%2:");
	return MapiObject::error(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE));
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

/**
 * We collect recipients as well as properties.
 */
bool MapiMessage::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
	static int ourTagList[] = {
		PidTagDisplayTo,
		PidTagDisplayCc,
		PidTagDisplayBcc,
		PidTagSenderEmailAddress,
		PidTagSenderSmtpAddress,
		PidTagSenderName,
		PidTagSenderSimpleDisplayName,
		PidTagOriginalSenderEmailAddress,
		PidTagOriginalSenderName,
		PidTagSentRepresentingEmailAddress,
		PidTagSentRepresentingName,
		PidTagSentRepresentingSimpleDisplayName,
		PidTagOriginalSentRepresentingEmailAddress,
		PidTagOriginalSentRepresentingName,
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
	if (!MapiObject::propertiesPull(tags, tagsAppended, pullAll)) {
		return false;
	}
	if (!recipientsPull()) {
		return false;
	}
	return true;
}

bool MapiMessage::propertiesPull()
{
	static bool tagsAppended = false;
	static QVector<int> tags;

	if (!propertiesPull(tags, tagsAppended, (DEBUG_MESSAGE_PROPERTIES) != 0)) {
		tagsAppended = true;
		return false;
	}
	tagsAppended = true;
	return true;
}

void MapiMessage::recipientPopulate(const char *phase, SRow &recipient, MapiRecipient &result)
{
	for (unsigned j = 0; j < recipient.cValues; j++) {
		MapiProperty property(recipient.lpProps[j]);
		QString tmp;

		// Note that the set of properties fetched here must be aligned
		// with those fetched in MapiConnector::resolveNames().
		switch (property.tag()) {
		case PidTag7BitDisplayName_string8:
		case PidTagDisplayName:
		case PidTagRecipientDisplayName:
			result.name = property.value().toString();
			tmp = mapiExtractEmail(property, "SMTP", true);
			if (isGoodEmailAddress(result.email) < isGoodEmailAddress(tmp)) {
				result.email = tmp;
			}
			break;
		case PidTagPrimarySmtpAddress:
			result.email = mapiExtractEmail(property, "SMTP");
			break;
		case UNDOCUMENTED_PR_EMAIL_UNICODE:
			debug() << "UNDOCUMENTED_PR_EMAIL_UNICODE" << property.value().toString();
			tmp = mapiExtractEmail(property, "SMTP");
			if (isGoodEmailAddress(result.email) < isGoodEmailAddress(tmp)) {
				result.email = tmp;
			}
			break;
		case PidTagRecipientTrackStatus:
			result.trackStatus = property.value().toInt();
			break;
		case PidTagRecipientFlags:
			result.flags = property.value().toInt();
			break;
		case PidTagRecipientType:
			// Mask off bits we don't want.
			result.setType((MapiRecipient::Type)(property.value().toUInt() & 0x3));
			break;
		case PidTagRecipientOrder:
			result.order = property.value().toInt();
			break;
		case PidTagDisplayType:
			result.setDisplayType((MapiRecipient::DisplayType)property.value().toUInt());
			break;
		case PidTagObjectType:
			result.setObjectType((MapiRecipient::ObjectType)property.value().toUInt());
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
#if (DEBUG_MESSAGE_PROPERTIES)
			debug() << "ignoring " << phase << " property:" << tagName(property.tag()) << property.value();
#else
			Q_UNUSED(phase);
#endif
			break;
		}
	}
}

/**
 * The recipients are pulled from multiple sources, but need to go through a 
 * resolution process to fix them up. The duplicates will disappear as part
 * of the resolution process. The logic looks like this:
 *
 *  1. Read the RecipientsTable.
 *
 *  2. Add in the contents of the DisplayTo tag (Exchange can return records
 *     from Step 1 with no name or email); ditto for CC and BCC.
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
 */bool MapiMessage::recipientsPull()
{
#if 0
// TEST -START-  Try an easier approach to find all the recipients
// Problem: this approach does not tell us about To/CC/BCC etc.

	struct ReadRecipientRow * recipientTable = 0x0;
	uint8_t recCount;
	if (MAPI_E_SSUCCESS != ReadRecipients(d(), 0, &recCount, &recipientTable))
	{
		error() << "cannot read recipient table" << mapiError();
		return false;
	}
	debug() << "number of recipients:" << recCount;
	for (int i = 0; i < recCount; i++) {
		struct RecipientRow &recipient = &recipientTable[i].RecipientRow;
		Recipient result;
		
		// Found this "intel" in MS-OXCDATA 2.8.3.1
		bool hasDisplayName = (recipient.RecipientFlags & 0x0010) != 0;
		if (hasDisplayName) {
			result.ame = QString::fromUtf8(recipient.DisplayName.lpszW);
		}
		bool hasEmailAddress = (recipient.RecipientFlags & 0x0008) != 0;
		if (hasEmailAddress) {
			result.email = QString::fromUtf8(recipient.EmailAddress.lpszW);
		}

		uint16_t type = recipient.RecipientFlags & 0x0007;
		switch (type) {
			case 0x1: // => X500DN
				// we need to resolve this recipient's data
				// TODO for now we just copy the user's account name to the recipientEmail
				recipientEmail = QString::fromLocal8Bit(recipient.X500DN.recipient_x500name);
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
#endif
	/**
	 * The list of tags used to fetch a recipient.
	 */
	SPropTagArray tableTags;

	// Start with a clean slate.
	m_recipients.clear();

	// Step 1. Add all the recipients from the actual table.
	SRowSet rowset;
	if (MAPI_E_SUCCESS != GetRecipientTable(&m_object, &rowset, &tableTags)) {
		error() << "cannot get recipient table:" << mapiError();
		return false;
	}

	for (unsigned i = 0; i < rowset.cRows; i++) {
		SRow &recipient = rowset.aRow[i];
		MapiRecipient result(MapiRecipient::To);

		recipientPopulate("recipient table", recipient, result);
		addUniqueRecipient("recipient table", result);
	}

	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	//
	// Step 2. Add the DisplayTo, CC and BCC and potential sender items.
	//
	// Astonishingly, when we call recipientsPull() later,
	// the results can contain entries with missing name (!)
	// and email values. Reading this property is a pathetic
	// workaround.
	MapiRecipient sender(MapiRecipient::Sender);
	MapiRecipient originalSender(MapiRecipient::Sender);
	MapiRecipient sentRepresenting(MapiRecipient::ReplyTo);
	MapiRecipient originalSentRepresenting(MapiRecipient::ReplyTo);

	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);

		switch (property.tag()) {
		case PidTagDisplayTo:
			foreach (QString name, property.value().toString().split(QChar::fromAscii(';'))) {
				MapiRecipient result(MapiRecipient::To);

				result.name = name.trimmed();
				result.email = mapiExtractEmail(result.name, "SMTP", true);
				addUniqueRecipient("displayTo", result);
			}
			break;
		case PidTagDisplayCc:
			foreach (QString name, property.value().toString().split(QChar::fromAscii(';'))) {
				MapiRecipient result(MapiRecipient::CC);

				result.name = name.trimmed();
				result.email = mapiExtractEmail(result.name, "SMTP", true);
				addUniqueRecipient("displayCC", result);
			}
			break;
		case PidTagDisplayBcc:
			foreach (QString name, property.value().toString().split(QChar::fromAscii(';'))) {
				MapiRecipient result(MapiRecipient::BCC);

				result.name = name.trimmed();
				result.email = mapiExtractEmail(result.name, "SMTP", true);
				addUniqueRecipient("displayBCC", result);
			}
			break;
		case PidTagSenderEmailAddress:
			sender.email = mapiExtractEmail(property, "EX");
			break;
		case PidTagSenderSmtpAddress:
			sender.email = mapiExtractEmail(property, "SMTP");
			break;
		CASE_PREFER_A_OVER_B(PidTagSenderName, PidTagSenderSimpleDisplayName, sender.name, property.value().toString());
			break;
		case PidTagOriginalSenderEmailAddress:
			originalSender.email = mapiExtractEmail(property, "SMTP");
			break;
		case PidTagOriginalSenderName:
			originalSender.name = property.value().toString();
			break;
		case PidTagSentRepresentingEmailAddress:
			sentRepresenting.email = mapiExtractEmail(property, "SMTP");
			break;
		CASE_PREFER_A_OVER_B(PidTagSentRepresentingName, PidTagSentRepresentingSimpleDisplayName, sentRepresenting.name, property.value().toString());
			break;
		case PidTagOriginalSentRepresentingEmailAddress:
			originalSentRepresenting.email = mapiExtractEmail(property, "SMTP");
			break;
		case PidTagOriginalSentRepresentingName:
			originalSentRepresenting.name = property.value().toString();
			break;
		}
	}

	// Step 3. Add the sent-on-behalf-of as well as the sender.
	if (!sender.name.isEmpty()) {
		// This might have come from a display name.
		QString tmp = mapiExtractEmail(sender.name, "SMTP", true);
		if (isGoodEmailAddress(sender.email) < isGoodEmailAddress(tmp)) {
			sender.email = tmp;
		}
		addUniqueRecipient("sender", sender);
	}
	if (!originalSender.name.isEmpty()) {
		addUniqueRecipient("originalSender", originalSender);
	}
	if (!sentRepresenting.name.isEmpty()) {
		// This might have come from a display name.
		QString tmp = mapiExtractEmail(sentRepresenting.name, "SMTP", true);
		if (isGoodEmailAddress(sentRepresenting.email) < isGoodEmailAddress(tmp)) {
			sentRepresenting.email = tmp;
		}
		addUniqueRecipient("replyTo", sentRepresenting);
	}
	if (!originalSentRepresenting.name.isEmpty()) {
		addUniqueRecipient("originalReplyTo", originalSentRepresenting);
	}

	// We have all the recipients; find any that need resolution.
	QList<int> needingResolution;
	for (int i = 0; i < m_recipients.size(); i++) {
		MapiRecipient &recipient = m_recipients[i];
		static QString perfectForm = QString::fromAscii("foo@foo");
		static unsigned perfect = isGoodEmailAddress(perfectForm);

		// If we find a missing/incomplete email, it needs resolution.
		if (isGoodEmailAddress(recipient.email) < perfect) {
			needingResolution << i;
		}
	}
	debug() << "recipients needing primary resolution:" << needingResolution.size() << "from a total:" << m_recipients.size();

	// Short-circuit exit.
	if (!needingResolution.size()) {
		return true;
	}

	// Primary resolution is to ask Exchange to resolve the names.
	struct PropertyTagArray_r *statuses = NULL;
	struct SRowSet *results = NULL;

	// Fill an array with the names we need to resolve. We will do a Unicode
	// lookup, so use UTF8.
	const char *names[needingResolution.size() + 1];
	unsigned j = 0;
	foreach (int i, needingResolution) {
		MapiRecipient &recipient = m_recipients[i];

		names[j] = string(recipient.name);
		j++;
	}
	names[j] = 0;

	// Server round trip here!
	static int recipientTagList[] = {
		PidTag7BitDisplayName_string8,
		PidTagDisplayName,
		PidTagRecipientDisplayName, 
		PidTagPrimarySmtpAddress,
		UNDOCUMENTED_PR_EMAIL_UNICODE,
		0x60010018,
		PidTagRecipientTrackStatus,
		PidTagRecipientFlags,
		PidTagRecipientType,
		PidTagRecipientOrder,
		PidTagDisplayType,
		PidTagObjectType,
		0 };
	static SPropTagArray recipientTags = {
		(sizeof(recipientTagList) / sizeof(recipientTagList[0])) - 1,
		(MAPITAGS *)recipientTagList };
	if (!m_connection->resolveNames(names, &recipientTags, &results, &statuses)) {
		return false;
	}
	if (results) {
		// Walk the returned results. Every request has a status, but
		// only resolved items also have a row of results.
		//
		// As we do the walk, we trim the needingResolution array so
		// that when we are done with this loop, it only contains
		// entries which need more work.
		for (unsigned i = 0, unresolveds = 0; i < statuses->cValues; i++) {
			MapiRecipient &to = m_recipients[needingResolution.at(unresolveds)];

			if (MAPI_RESOLVED == statuses->aulPropTag[i]) {
				struct SRow &recipient = results->aRow[i - unresolveds];
				MapiRecipient result(MapiRecipient::To);

				recipientPopulate("resolution", recipient, result);
				// A resolved value is better than an unresolved one.
				if (!result.name.isEmpty()) {
					to.name = result.name;
				}
				if (isGoodEmailAddress(to.email) < isGoodEmailAddress(result.email)) {
					to.email = result.email;
				}
				needingResolution.removeAt(unresolveds);
			} else {
				unresolveds++;
			}
		}
	}
	MAPIFreeBuffer(results);
	MAPIFreeBuffer(statuses);
	debug() << "recipients needing secondary resolution:" << needingResolution.size();

	// Secondary resolution is to remove entries which have the the same 
	// email. But we must take care since we'll still have unresolved 
	// entries with empty email values (which would collide).
	QMap<QString, MapiRecipient> uniqueResolvedRecipients;
	for (int i = 0; i < m_recipients.size(); i++) {
		MapiRecipient &recipient = m_recipients[i];

		if (!recipient.email.isEmpty()) {
			QMap<QString, MapiRecipient>::const_iterator i = uniqueResolvedRecipients.constFind(recipient.email);
			if (i != uniqueResolvedRecipients.constEnd()) {
				// If we have duplicates, keep the one with the longer 
				// name in the hopes that it will be the more descriptive.
				if (recipient.name.length() < i.value().name.length()) {
					continue;
				}
				
				// Don't allow different types to collide to preserve ReplyTo
				// distinctiveness.
				if (recipient.type() != i.value().type()) {
					uniqueResolvedRecipients.insertMulti(recipient.email, recipient);
					continue;
				}
			}
			uniqueResolvedRecipients.insert(recipient.email, recipient);

			// If the name we just inserted matches one with an empty
			// email, we can kill the latter.
			for (int j = 0; j < needingResolution.size(); j++)
			{
				if (recipient.name == m_recipients[needingResolution.at(j)].name) {
					needingResolution.removeAt(j);
					j--;
				}
			}
		}
	}
	debug() << "recipients needing tertiary resolution:" << needingResolution.size();

	// Tertiary resolution.
	//
	// Add items needing more work. They should have an empty email. Just 
	// suck it up accepting the fact they'll be "duplicates" by key, but 
	// not by name. And delete any which have no name either.
	for (int i = 0; i < needingResolution.size(); i++)
	{
		MapiRecipient &recipient = m_recipients[needingResolution.at(i)];

		if (recipient.name.isEmpty()) {
			// Grrrr...
			continue;
		}

		// Set the email to be the same as the name. It cannot be any
		// worse, right?
		recipient.email = recipient.name;
		uniqueResolvedRecipients.insertMulti(recipient.email, recipient);
	}

	// Finally, recreate the sanitised list.
	m_recipients.clear();
	foreach (MapiRecipient recipient, uniqueResolvedRecipients) {
		m_recipients.append(recipient);
	error() << "recipient name:" << recipient.toString();
	}
	debug() << "recipients after resolution:" << m_recipients.size();
	return true;
}

const QList<MapiRecipient> &MapiMessage::recipients()
{
	return m_recipients;
}

MapiObject::MapiObject(MapiConnector2 *connection, const char *tallocName, mapi_id_t id) :
	TallocContext(tallocName),
	m_connection(connection),
	m_id(id),
	m_properties(0),
	m_propertyCount(0),
	m_ourTagList(0)
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

bool MapiObject::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
	if (!tagsAppended || !m_ourTagList) {
		m_ourTagList = array<int>(tags.size() + 1);
		if (!m_ourTagList) {
			error() << "cannot allocate tags:" << tags.size() << mapiError();
			return false;
		}
		m_ourTags.cValues = tags.size();
		m_ourTags.aulPropTag = (MAPITAGS *)m_ourTagList;
		unsigned i = 0;
		foreach (int tag, tags) {
			m_ourTagList[i++] = tag;
		}
		m_ourTagList[i] = 0;
	}

	m_properties = 0;
	m_propertyCount = 0;
	if (pullAll) {
		return MapiObject::propertiesPull();
	}
	if (MAPI_E_SUCCESS != GetProps(&m_object, MAPI_UNICODE, &m_ourTags, &m_properties, &m_propertyCount)) {
		error() << "cannot pull properties:" << mapiError();
		return false;
	}
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
	m_properties = array<SPropValue>(mapiProperties.cValues);
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
	char *copy = string(data);

	if (!copy) {
		error() << "cannot talloc:" << data;
		return false;
	}

	return propertyWrite(tag, copy, idempotent);
}

bool MapiObject::propertyWrite(int tag, QDateTime &data, bool idempotent)
{
	FILETIME *copy = allocate<FILETIME>();

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
	const char *str;

	// Some properties share ids...
	switch (tag) {
	case PidTagAttachDataObject:
		str = "PidTagAttachDataObject";
		break;
	case PidTagAttachDataObject_Error:
		str = "PidTagAttachDataObject_Error";
		break;
	default:
		str = get_proptag_name(tag);
	}

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

	// It seams that the QByteArray returned by QString::toUtf8() are not backed up.
	// calling toUtf8() several times will destroy the data. Therefore I copy them
	// all to temporary objects in order to preserve them throughout the whole method
	QByteArray baProfile(profile.toUtf8());
	QByteArray baUser(username.toUtf8());
	QByteArray baPass(password.toUtf8());
	QByteArray baDomain(domain.toUtf8());
	QByteArray baServer(server.toUtf8());
	
	const char *profile8 = baProfile.constData();
	qDebug() << "New profile is:"<<profile8;

	if (MAPI_E_SUCCESS != CreateProfile(m_context, profile8, baUser.constData(), baPass.constData(), 0)) {
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
		if ((m_property.ulPropTag & 0xFFFF) == PT_ERROR) {
			return QString::number((unsigned)m_property.value.err, 16);
		} else {
			return tmp.toString();
		}
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

QString MapiRecipient::toString() const
{
	static QString format = QString::fromAscii("%1 <%2>, %3, %4, %5");

	return format.arg(name).arg(email).arg(typeString()).arg(objectTypeString()).arg(displayTypeString());
}

QString MapiRecipient::displayTypeString() const
{
	switch (m_displayType)
	{
	case DtMailuser:
		return QString::fromAscii("Messaging user");
	case DtDistlist:
		return QString::fromAscii("Distribution list");
	case DtForum:
		return QString::fromAscii("Forum");
	case DtAgent:
		return QString::fromAscii("Automated agent");
	case DtOrganization:
		return QString::fromAscii("Group alias");
	case DtPrivateDistlist:
		return QString::fromAscii("Private distribution list");
	case DtRemoteMailuser:
		return QString::fromAscii("Foreign/remote messaging user");
	case DtRoom:
		return QString::fromAscii("Room");
	case DtEquipment:
		return QString::fromAscii("Equipment");
	case DtSecurityGroup:
		return QString::fromAscii("Security group");
	default:
		return QString::fromAscii("MAPI_0x%1 display type").arg(m_objectType, 0, 16);
	}
}

QString MapiRecipient::objectTypeString() const
{
	switch (m_objectType)
	{
	case OtStore:
		return QString::fromAscii("Message store");
	case OtAddrbook:
		return QString::fromAscii("Address book");
	case OtFolder:
		return QString::fromAscii("Folder");
	case OtABcont:
		return QString::fromAscii("Address book container");
	case OtMessage:
		return QString::fromAscii("Message");
	case OtMailuser:
		return QString::fromAscii("Messaging user");
	case OtAttach:
		return QString::fromAscii("Message attachment");
	case OtDistlist:
		return QString::fromAscii("Distribution list");
	case OtProfsect:
		return QString::fromAscii("Profile section");
	case OtStatus:
		return QString::fromAscii("Status");
	case OtSession:
		return QString::fromAscii("Session");
	case OtForminfo:
		return QString::fromAscii("Form");
	default:
		return QString::fromAscii("MAPI_0x%1 object type").arg(m_displayType, 0, 16);
	}
}

QString MapiRecipient::typeString() const
{
	switch (m_type)
	{
	case Sender:
		return QString::fromAscii("Sender");
	case To:
		return QString::fromAscii("To");
	case CC:
		return QString::fromAscii("CC");
	case BCC:
		return QString::fromAscii("BCC");
	case ReplyTo:
		return QString::fromAscii("ReplyTo");
	default:
		return QString::fromAscii("MAPI_0x%1 type").arg(m_type, 0, 16);
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
	if (!m_ctx) {
		qCritical() << name << "talloc_named failed";
	}
}

TallocContext::~TallocContext()
{
	talloc_free(m_ctx);
}

template <class T>
T *TallocContext::allocate()
{
	return talloc(m_ctx, T);
}

template <class T>
T *TallocContext::array(unsigned size)
{
	return talloc_array(m_ctx, T, size);
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

char *TallocContext::string(const QString &original)
{
	return talloc_strdup(m_ctx, original.toUtf8().data());
}
