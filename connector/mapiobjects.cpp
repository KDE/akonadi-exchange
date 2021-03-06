/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011-13 Robert Gruber <rgruber@users.sourceforge.net>, Shaheed Haque
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

#include <QAbstractSocket>
#include <QDebug>
#include <QStringList>
#include <QDir>
#include <QMessageBox>
#include <QRegExp>
#include <QVariant>
#include <QSocketNotifier>
#include <QTextCodec>
#include <KLocale>
#include <kpimutils/email.h>
#include "mapiobjects.h"

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

#ifndef DEBUG_RECIPIENTS
#define DEBUG_RECIPIENTS 0
#endif

#ifndef DEBUG_NOTIFICATIONS
#define DEBUG_NOTIFICATIONS 0
#endif

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

MapiFolder::MapiFolder(MapiConnector2 *connection, const char *tallocName, const MapiId &id) :
    MapiObject(connection, tallocName, id)
{
    mapi_object_init(&m_contents);
    // A temporary name.
    name = id.toString();
}

MapiFolder::~MapiFolder()
{
    mapi_object_release(&m_contents);
}

QDebug MapiFolder::debug() const
{
    static QString prefix = QString::fromAscii("MapiFolder: %1:");
    return MapiObject::debug(prefix.arg(m_id.toString()));
}

QDebug MapiFolder::error() const
{
    static QString prefix = QString::fromAscii("MapiFolder: %1:");
    return MapiObject::error(prefix.arg(m_id.toString()));
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
            mapi_id_t fid = 0;
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
            MapiId folderId(m_id, fid);
            MapiFolder *data = new MapiFolder(m_connection, "MapiFolder::childrenPull", folderId);
            data->name = name;
            children.append(data);
        }
    }
    return true;
}

bool MapiFolder::childrenPull(QList<MapiItem *> &children)
{
    // Retrieve folder's content table
    if (MAPI_E_SUCCESS != GetContentsTable(&m_object, &m_contents, TableFlags_UseUnicode, NULL)) {
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
            mapi_id_t id = 0;
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
            MapiId itemId(m_id, id);
            MapiItem *data = new MapiItem(itemId, name, modified);
            children.append(data);
            //TODO Just for debugging (in case the content list ist very long)
            //if (i >= 10) break;
        }
    }
    return true;
}

bool MapiFolder::open()
{
    if (MAPI_E_SUCCESS != OpenFolder(m_connection->store(m_id), m_id.second, &m_object)) {
        error() << "cannot open folder" << m_id << mapiError();
        return false;
    }
    return true;
}

MapiItem::MapiItem(const MapiId &id, QString &name, QDateTime &modified) :
    m_id(id),
    m_name(name),
    m_modified(modified)
{
}

const MapiId &MapiItem::id() const
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

MapiMessage::MapiMessage(MapiConnector2 *connection, const char *tallocName, const MapiId &id) :
    MapiObject(connection, tallocName, id)
{
}

void MapiMessage::addUniqueRecipient(const char *source, MapiRecipient &candidate)
{
#if DEBUG_RECIPIENTS
    debug() << "candidate address:" << source << candidate.toString();
#else
    Q_UNUSED(source)
#endif
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
#if DEBUG_RECIPIENTS
    debug() << "add new address:" << source << candidate.toString();
#endif
    m_recipients.append(candidate);
}

QDebug MapiMessage::debug() const
{
    static QString prefix = QString::fromAscii("MapiMessage: %1:");
    return MapiObject::debug(prefix.arg(m_id.toString()));
}

QDebug MapiMessage::error() const
{
    static QString prefix = QString::fromAscii("MapiMessage: %1:");
    return MapiObject::error(prefix.arg(m_id.toString()));
}

bool MapiMessage::open()
{
    if (MAPI_E_SUCCESS != OpenMessage(m_connection->store(m_id), m_id.first, m_id.second, &m_object, 0x0)) {
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
    static unsigned ourTagList[] = {
        PidTagDisplayTo,
        PidTagDisplayCc,
        PidTagDisplayBcc,
        PidTagSenderEmailAddress,
        PidTagSenderSmtpAddress,
        PidTagSenderName,
        PidTagOriginalSenderEmailAddress,
        PidTagOriginalSenderName,
        PidTagSentRepresentingEmailAddress,
        PidTagSentRepresentingName,
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
        case PidTagDisplayName_string8:
        case PidTagDisplayName:
        case PidTagRecipientDisplayName:
            result.name = property.value().toString();
            tmp = mapiExtractEmail(property, "SMTP", true);
            if (isGoodEmailAddress(result.email) < isGoodEmailAddress(tmp)) {
                result.email = tmp;
            }
            break;
        case PidTagSmtpAddress:
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
 */
bool MapiMessage::recipientsPull()
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
        case PidTagSenderName:
            sender.name = property.value().toString();
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
        case PidTagSentRepresentingName:
            sentRepresenting.name = property.value().toString();
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
#if DEBUG_RECIPIENTS
    debug() << "recipients needing primary resolution:" << needingResolution.size() << "from a total:" << m_recipients.size();
#endif

    // Short-circuit exit.
    if (!needingResolution.size()) {
        return true;
    }

    // Primary resolution is to ask Exchange to resolve the names.
    struct PropertyTagArray_r *statuses = NULL;
    SRowSet *results = NULL;

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
        PidTagDisplayName_string8,
        PidTagDisplayName,
        PidTagRecipientDisplayName, 
        PidTagSmtpAddress,
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
#if DEBUG_RECIPIENTS
    debug() << "recipients needing secondary resolution:" << needingResolution.size();
#endif

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
#if DEBUG_RECIPIENTS
    debug() << "recipients needing tertiary resolution:" << needingResolution.size();
#endif

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
#if DEBUG_RECIPIENTS
        debug() << "recipient name:" << recipient.toString();
#endif
    }
#if DEBUG_RECIPIENTS
    debug() << "recipients after resolution:" << m_recipients.size();
#endif
    return true;
}

const QList<MapiRecipient> &MapiMessage::recipients()
{
    return m_recipients;
}

bool MapiMessage::streamRead(mapi_object_t *parent, int tag, QByteArray &bytes)
{
    mapi_object_t stream;
    unsigned dataSize;
    unsigned offset;
    uint16_t readSize;

    mapi_object_init(&stream);
    if (MAPI_E_SUCCESS != OpenStream(parent, (MAPITAGS)tag, OpenStream_ReadOnly, &stream)) {
        error() << "cannot open stream:" << tagName(tag) << mapiError();
        mapi_object_release(&stream);
        return false;
    }
    if (MAPI_E_SUCCESS != GetStreamSize(&stream, &dataSize)) {
        error() << "cannot get stream size:" << tagName(tag) << mapiError();
        mapi_object_release(&stream);
        return false;
    }
    bytes.reserve(dataSize);
    offset = 0;
    do {
        if (MAPI_E_SUCCESS != ReadStream(&stream, (uchar *)bytes.data() + offset, 0x1000, &readSize)) {
            error() << "cannot read stream:" << tagName(tag) << mapiError();
            mapi_object_release(&stream);
            return false;
        }
        offset += readSize;
    } while (readSize && (offset < dataSize));
    bytes.resize(dataSize);
    mapi_object_release(&stream);
    return true;
}

/**
 * See the codepage2codec map below.
 */
const unsigned MapiMessage::CODEPAGE_UTF16 = 1200;

bool MapiMessage::streamRead(mapi_object_t *parent, int tag, unsigned codepage, QString &string)
{
    // Map QTextCodec names to Microsoft Code Pages
    typedef struct
    {
        unsigned codepage;
        const char *codec;
    } codepage2codec;

    static codepage2codec map[] =
    {
        { 10000,	"Apple Roman" },
        { 950,		"Big5" },
        { 950,		"Big5-HKSCS" },
        { 949,		"CP949" },
        { 20932,	"EUC-JP" },
        { 51949,	"EUC-KR" },
        { 54936,	"GB18030-0" },
        { 850,		"IBM 850" },
        { 866,		"IBM 866" },
        { 874,		"IBM 874" },
        { 50220,	"ISO 2022-JP" },
        { 28591,	"ISO 8859-1" },
        { 28592,	"ISO 8859-2" },
        { 28593,	"ISO 8859-3" },
        { 28594,	"ISO 8859-4" },
        { 28595,	"ISO 8859-5" },
        { 28596,	"ISO 8859-6" },
        { 28597,	"ISO 8859-7" },
        { 28598,	"ISO 8859-8" },
        { 28599,	"ISO 8859-9" },
        { 28600,	"ISO 8859-10" },
        { 28603,	"ISO 8859-13" },
        { 28604,	"ISO 8859-14" },
        { 28605,	"ISO 8859-15" },
        { 28606,	"ISO 8859-16" },
        { 57003,	"Iscii-Bng" },
        { 57002,	"Iscii-Dev" },
        { 57010,	"Iscii-Gjr" },
        { 57008,	"Iscii-Knd" },
        { 57009,	"Iscii-Mlm" },
        { 57007,	"Iscii-Ori" },
        { 57011,	"Iscii-Pnj" },
        { 57005,	"Iscii-Tlg" },
        { 57004,	"Iscii-Tml" },
        { 50222,	"JIS X 0201" },
        { 20932,	"JIS X 0208" },
        { 20866,	"KOI8-R" },
        { 21866,	"KOI8-U" },
        //{,		"MuleLao-1" },
        //{,		"ROMAN8" },
        { 932,		"Shift-JIS" },
        { 874,		"TIS-620" },
        { 57004,	"TSCII" },
        { 65001,	"UTF-8" },
        { CODEPAGE_UTF16,"UTF-16" },
        { 1201,		"UTF-16BE" },
        { 1200,		"UTF-16LE" },
        { 12000,	"UTF-32" },
        { 12001,	"UTF-32BE" },
        { 12000,	"UTF-32LE" },
        { 1250,		"Windows-1250" },
        { 1251,		"Windows-1251" },
        { 1252,		"Windows-1252" },
        { 1253,		"Windows-1253" },
        { 1254,		"Windows-1254" },
        { 1255,		"Windows-1255" },
        { 1256,		"Windows-1256" },
        { 1257,		"Windows-1257" },
        { 1258,		"Windows-1258" },
        //{,		"WINSAMI2" },
        { 0, 0 }
    };
    QByteArray bytes;
    codepage2codec *entry = &map[0];

    while (entry->codepage && entry->codepage != codepage) {
        entry++;
    }
    if (!entry->codec) {
        error() << "codec name not found for codepage:" << codepage;
        return false;
    }
    if (!streamRead(parent, tag, bytes)) {
        return false;
    }
    QTextCodec *codec = QTextCodec::codecForName(entry->codec);
    string = codec->toUnicode(bytes);
    return true;
}

MapiObject::MapiObject(MapiConnector2 *connection, const char *tallocName, const MapiId &id) :
    TallocContext(tallocName),
    m_connection(connection),
    m_id(id),
    m_properties(0),
    m_propertyCount(0),
    m_cachedNames(0),
    m_cachedNamedTags(0),
    m_listenerId(0)
{
    mapi_object_init(&m_object);
    m_cachedTags.cValues = 0;
    m_cachedTags.aulPropTag = 0;
}

MapiObject::~MapiObject()
{
    mapi_object_release(&m_object);
}

mapi_object_t *MapiObject::d() const
{
    return &m_object;
}

const MapiId &MapiObject::id() const
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
    bool usingNamedProperties = false;

    // If the user tells us the tags he has given us have not previously been
    // seen, or the cache is empty, fill it.
    if (!tagsAppended || !m_cachedTags.aulPropTag) {
        // First create the list of tags.
        m_cachedTags.aulPropTag = (MAPITAGS *)array<int>(tags.size() + 1);
        if (!m_cachedTags.aulPropTag) {
            error() << "cannot allocate tags:" << tags.size() << mapiError();
            return false;
        }

        // Now create the array that MAPI will use.
        m_cachedTags.cValues = tags.size();
        unsigned i = 0;
        foreach (int tag, tags) {
            m_cachedTags.aulPropTag[i++] = (MAPITAGS)tag;
            usingNamedProperties |= ((tag & 0x80000000) != 0);
        }
        m_cachedTags.aulPropTag[i] = (MAPITAGS)0;

        // Now create the named property translation, if needed.
        if (usingNamedProperties) {
            m_cachedNames = mapi_nameid_new(ctx());
            if (!m_cachedNames) {
                error() << "Cannot create named property context" << mapiError();
                return false;
            }
            m_cachedNamedTags = talloc_zero(ctx(), struct SPropTagArray);
            if (!m_cachedNamedTags) {
                error() << "Cannot create named property array" << mapiError();
                return false;
            }
            MAPISTATUS status = mapi_nameid_lookup_SPropTagArray(m_cachedNames, &m_cachedTags);
            if ((MAPI_E_NOT_FOUND != status) && (MAPI_E_SUCCESS != status)) {
                error() << "Cannot find named properties" << mapiError();
                return false;
            }
            if (MAPI_E_SUCCESS != mapi_nameid_GetIDsFromNames(m_cachedNames, &m_object, m_cachedNamedTags)) {
                error() << "Cannot find named property ids" << mapiError();
                return false;
            }
        } else {
            m_cachedNames = 0;
        }
    }

    m_properties = 0;
    m_propertyCount = 0;
    if (pullAll) {
        return MapiObject::propertiesPull();
    }

    // Map the tags before the call, and unmap them afterwards.
    if (usingNamedProperties) {
        if (MAPI_E_SUCCESS != mapi_nameid_map_SPropTagArray(m_cachedNames, &m_cachedTags, m_cachedNamedTags)) {
            error() << "Cannot map named properties" << mapiError();
            return false;
        }
    }
    if (MAPI_E_SUCCESS != GetProps(&m_object, MAPI_UNICODE | MAPI_PROPS_SKIP_NAMEDID_CHECK, &m_cachedTags, &m_properties, &m_propertyCount)) {
        error() << "cannot pull properties:" << mapiError();
        if (usingNamedProperties) {
            if (MAPI_E_SUCCESS != mapi_nameid_unmap_SPropTagArray(m_cachedNames, &m_cachedTags)) {
                error() << "Cannot unmap named properties" << mapiError();
                return false;
            }
        }
        return false;
    }
    if (usingNamedProperties) {
        if (MAPI_E_SUCCESS != mapi_nameid_unmap_SPropTagArray(m_cachedNames, &m_cachedTags)) {
            error() << "Cannot unmap named properties" << mapiError();
            return false;
        }

        // When we unmap the property tags in the value array, preserve the type
        // portion since that's how we get to know about errors.
        QVector<int> types(m_propertyCount);
        for (unsigned i = 0; i < m_propertyCount; i++) {
            types[i] = m_properties[i].ulPropTag & 0xFFFF;
        }
        if (MAPI_E_SUCCESS != mapi_nameid_unmap_SPropValue(m_cachedNames, m_properties, m_propertyCount)) {
            error() << "Cannot unmap named values" << mapiError();
            return false;
        }
        for (unsigned i = 0; i < m_propertyCount; i++) {
            m_properties[i].ulPropTag = (MAPITAGS)(m_properties[i].ulPropTag & (0xFFFF0000 | types[i]));
        }
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

bool MapiObject::subscribe()
{
    if (MAPI_E_SUCCESS != Subscribe(&m_object, &m_listenerId, -1, false, 0, this)) {
        error() << "cannot subscribe listener" << mapiError();
        return false;
    }
    return true;
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
    return toString(m_displayType);
}

QString MapiRecipient::toString(DisplayType type, unsigned pluralForm)
{
    switch (type)
    {
    case DtMailuser:
        return i18np("User", "Users", pluralForm);
    case DtDistlist:
        return i18np("List", "Lists", pluralForm);
    case DtForum:
        return i18np("Forum", "Forums", pluralForm);
    case DtAgent:
        return i18np("Agent", "Agents", pluralForm);
    case DtOrganization:
        return i18np("Group Alias", "Group Aliases", pluralForm);
    case DtPrivateDistlist:
        return i18np("Private List", "Private Lists", pluralForm);
    case DtRemoteMailuser:
        return i18np("External User", "External Users", pluralForm);
    case DtRoom:
        return i18np("Room", "Rooms", pluralForm);
    case DtEquipment:
        return i18np("Equipment", "Equipment", pluralForm);
    case DtSecurityGroup:
        return i18np("Security Group", "Security Groups", pluralForm);
    default:
        return i18n("DisplayType%1", type);
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
