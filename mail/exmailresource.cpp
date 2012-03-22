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
#include <akonadi/kmime/messageflags.h>
#include <akonadi/kmime/messageparts.h>
#include <kmime/kmime_message.h>
#include <kmime/kmime_util.h>
#include <kpimutils/email.h>

#include "mapiconnector2.h"
#include "profiledialog.h"

/**
 * Set this to 1 to pull all the properties, e.g. to see what a server has
 * available.
 */
#ifndef DEBUG_NOTE_PROPERTIES
#define DEBUG_NOTE_PROPERTIES 0
#endif

#define GET_SUBJECTS_FOR_EMBEDDED_MSGS 0

using namespace Akonadi;

/**
 * An Email. Note that the MAPI service offered by Exchange does not give us
 * access to the raw message that might have worked its way across the Internet.
 * 
 * For example, where the top level has multipart/alternatives for HTML and
 * text, it'll only give us the former. Therefore, the model of KMime::Content
 * we'll be able to provide will not, in general match the IMAP interface.
 */
class MapiNote : public MapiMessage, public KMime::Message
{
public:
    MapiNote(MapiConnector2 *connection, const char *tallocName, MapiId &id);

    virtual ~MapiNote();

    /**
     * Fetch all email properties.
     */
    virtual bool propertiesPull();

    /**
     * Update a note item.
     */
    virtual bool propertiesPush();

protected:
    virtual QDebug debug() const;
    virtual QDebug error() const;

    /**
     * Populate the object with property contents.
     */
    bool preparePayload();

    /**
     * Fetch email properties.
     */
    virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);

    mapi_object_t m_attachments;
    mapi_object_t m_attachment;
};

/**
 * An embedded Email. This is exactly the same as @ref MapiNote except that a
 * different table contains all the properties (and we have a parent attachment
 * object whose content we embody).
 */
class MapiEmbeddedNote: public MapiNote
{
public:
    MapiEmbeddedNote(MapiConnector2 *connector, const char *tallocName, MapiId &id, mapi_object_t *parentAttachment);

    bool open();

#if (GET_SUBJECTS_FOR_EMBEDDED_MSGS)
    /**
     * Fetch all embedded email properties.
     */
    virtual bool propertiesPull();

protected:

    /**
     * Fetch embedded email properties. This is different than for MapiNote
     * because that uses PidTagTransportMessageHeaders for maximum fidelity,
     * but embedded messages don't come with such luxuries.
     */
    virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);
#endif

private:
    mapi_object_t *m_parentAttachment;
};

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

const QString ExMailResource::profile()
{
    // First select who to log in as.
    return Settings::self()->profileName();
}

void ExMailResource::retrieveCollections()
{
    Collection::List collections;

    setName(i18n("Exchange Mail for %1", profile()));
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
//#if (DEBUG_NOTE_PROPERTIES)
    while (items.size() > 3) {
        items.removeLast();
    }
//#endif
    itemsRetrievedIncremental(items, deletedItems);
}

bool ExMailResource::retrieveItem(const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts)
{
    Q_UNUSED(parts);

    MapiNote *message = fetchItem<MapiNote>(itemOrig);
    if (!message) {
        return false;
    }
    KMime::Message::Ptr ptr(message);

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
    // TODO add further message properties.
    //item.setModificationTime(message->modified);
*/
    item.setPayload<KMime::Message::Ptr>(ptr);

    // Notify Akonadi about the new data.
    itemRetrieved(item);
    return true;
}

void ExMailResource::aboutToQuit()
{
  // TODO: any cleanup you need to do while there is still an active
  // event loop. The resource will terminate after this method returns
}

void ExMailResource::configure(WId windowId)
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

MapiEmbeddedNote::MapiEmbeddedNote(MapiConnector2 *connector, const char *tallocName, MapiId &id, mapi_object_t *parentAttachment) :
    MapiNote(connector, tallocName, id),
    m_parentAttachment(parentAttachment)
{
}

bool MapiEmbeddedNote::open()
{
    if (MAPI_E_SUCCESS != OpenEmbeddedMessage(m_parentAttachment, &m_object, MAPI_READONLY)) {
        error() << "cannot open embedded message, error:" << mapiError();
        return false;
    }
    return true;
}

#if (GET_SUBJECTS_FOR_EMBEDDED_MSGS)
bool MapiEmbeddedNote::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
    /**
     * The list of tags used to fetch an embedded Note, based on [MS-OXCMSG].
     */
    static int ourTagList[] = {
        PidTagSubject,
        PidTagCreationTime,
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
    if (!MapiNote::propertiesPull(tags, tagsAppended, pullAll)) {
        return false;
    }
    if (!preparePayload()) {
        return false;
    }
    return true;
}

bool MapiEmbeddedNote::propertiesPull()
{
    static bool tagsAppended = false;
    static QVector<int> tags;

    if (!propertiesPull(tags, tagsAppended, (DEBUG_NOTE_PROPERTIES) != 0)) {
        tagsAppended = true;
        return false;
    }
    tagsAppended = true;
    return true;
}
#endif

MapiNote::MapiNote(MapiConnector2 *connector, const char *tallocName, MapiId &id) :
    MapiMessage(connector, tallocName, id),
    KMime::Message()
{
    mapi_object_init(&m_attachments);
    mapi_object_init(&m_attachment);
}
 
MapiNote::~MapiNote()
{
    mapi_object_release(&m_attachment);
    mapi_object_release(&m_attachments);
}

QDebug MapiNote::debug() const
{
    static QString prefix = QString::fromAscii("MapiNote: %1:");
    return MapiObject::debug(prefix.arg(m_id.toString()));
}

QDebug MapiNote::error() const
{
    static QString prefix = QString::fromAscii("MapiNote: %1");
    return MapiObject::error(prefix.arg(m_id.toString()));
}

/**
 * Create the "raw source" as well as all the properties we need.
 * 
 * The switch statement at the heart of this routine must be kept synchronised
 * with @ref noteTagList.
 *
 * @return false on error.
 */
bool MapiNote::preparePayload()
{
    unsigned index;
    unsigned codepage = 0;
    QString textBody;
    QString htmlBody;
    bool hasAttachments = false;

    // First set the header content, and parse what we can from it. Note
    // that the message headers we are given:
    //
    //	- Start with some kind of descriptive string.
    //	- End each line with CRLF.
    //	- Have all subcontent headers and boundaries.
    //
    // like this:
    //
    // Microsoft Mail Internet Headers Version 2.0^M
    // Content-Type: multipart/alternative;^M
    // boundary="----_=_NextPart_001_01CC9A5A.689896D8"^M
    // Subject: ...^M
    // ...
    // Return-Path: build-guru@example.com^M
    // ^M
    // ------_=_NextPart_001_01CC9A5A.689896D8^M
    // Content-Type: text/plain;^M
    //         charset="US-ASCII"^M
    // Content-Transfer-Encoding: quoted-printable^M
    // ^M
    // ------_=_NextPart_001_01CC9A5A.689896D8^M
    // Content-Type: text/html;^M
    //         charset="US-ASCII"^M
    //
    // However, the parse() routine will return the last value only. For
    // example, the above will result in a contentType of "text/html" 
    // instead of "multipart/alternative". For all these reasons, we need 
    // a fixed-up version to work with.
    if (UINT_MAX > (index = propertyFind(PidTagTransportMessageHeaders))) {
        QString header = propertyAt(index).toString().prepend(QString::fromAscii("X-Parsed-By: "));

        // Fixup the header.
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
            case '\t':
            case ' ':
                // Unfold?
                if (lastChWasNl) {
                    j--;
                    break;
                }
                // Copy anything else.
                header[j] = ch;
                j++;
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

        // Set up all the headers. For unknown reasons, parse() gets
        // the Content-Type wrong, so we have to fix that up by hand.
        setHead(header.toUtf8());
        parse();
        QByteArray tmp = KMime::extractHeader(head(), "Content-Type");
        contentType()->from7BitString(tmp);
    }

    // Walk through the properties and extract the values of interest. The
    // properties here should be aligned with the list pulled above.
    bool textStream = false;
    bool htmlStream = false;
    for (unsigned i = 0; i < m_propertyCount; i++) {
        MapiProperty property(m_properties[i]);

        switch (property.tag()) {
        case PidTagMessageClass:
            // Sanity check the message class.
            if ((QLatin1String("IPM.Note") != property.value().toString()) &&
                (QLatin1String("Remote.IPM.Note") != property.value().toString())){
                error() << "retrieved item is not an email or a header:" << property.value().toString();
                return false;
            }
            break;
        case PidTagMessageCodepage:
            codepage = property.value().toUInt();
            break;
        case PidTagMessageFlags:
            hasAttachments = (property.value().toUInt() & MSGFLAG_HASATTACH) != 0;
            break;
#if (GET_SUBJECTS_FOR_EMBEDDED_MSGS)
        case PidTagSubject:
            subject()->fromUnicodeString(property.value().toString(), "utf-8");
            break;
#endif
        case PidTagBody:
            textBody = property.value().toString();
            break;
        case PidTagHtml:
            htmlBody = property.value().toString();
            break;
        case PidTagTransportMessageHeaders:
            break;
#if (GET_SUBJECTS_FOR_EMBEDDED_MSGS)
        case PidTagCreationTime:
            date()->setDateTime(KDateTime(property.value().toDateTime()));
            break;
#endif
        default:
            // Handle oversize objects.
            if (MAPI_E_NOT_ENOUGH_MEMORY == property.value().toInt()) {
                switch (property.tag()) {
                case PidTagBody_Error:
                    textStream = true;
                    break;
                case PidTagHtml_Error:
                    htmlStream = true;
                    break;
                default:
                    error() << "missing oversize support:" << tagName(property.tag());
                    break;
                }

                // Carry on with next property...
                break;
            }
//#if (DEBUG_NOTE_PROPERTIES)
            debug() << "ignoring note property:" << tagName(property.tag()) << property.toString();
//#endif
            break;
        }
    }

    foreach (MapiRecipient item, MapiMessage::recipients()) {
        switch (item.type()) {
        case MapiRecipient::Sender:
            from()->addAddress(item.email.toUtf8(), item.name);
            break;
        case MapiRecipient::To:
            to()->addAddress(item.email.toUtf8(), item.name);
            break;
        case MapiRecipient::CC:
            cc()->addAddress(item.email.toUtf8(), item.name);
            break;
        case MapiRecipient::BCC:
            bcc()->addAddress(item.email.toUtf8(), item.name);
            break;
        case MapiRecipient::ReplyTo:
            replyTo()->addAddress(item.email.toUtf8(), item.name);
            break;
        }
    }

    // We get the PidTagBody as Unicode in any event, but we also now know
    // the codepage for PidTagHtml.
    if (textStream && !streamRead(&m_object, PidTagBody, CODEPAGE_UTF16, textBody)) {
        return false;
    }
    if (htmlStream && !streamRead(&m_object, PidTagHtml, codepage, htmlBody)) {
        return false;
    }
    error() << "text size:" << textBody.size() << "html size:" << htmlBody.size() << "attachments:" << hasAttachments << "mimeType:" << contentType()->mimeType() << "isEmbedded:" << dynamic_cast<MapiEmbeddedNote*>(this);

    // If we don't have a Content-Type, then one will be automatically added.
    // Unfortunately, when that happens, we seem to get some bogus, empty
    // body parts added. So, let's make sure we have something whenever we can.
    if (!contentType()->mimeType().size()) {
        if (hasAttachments) {
            contentType()->setMimeType("multipart/mixed");
            contentType()->setBoundary(KMime::multiPartBoundary());
        } else if (!textBody.isEmpty() && !htmlBody.isEmpty()) {
            contentType()->setMimeType("multipart/alternative");
            contentType()->setBoundary(KMime::multiPartBoundary());
        } else if (!textBody.isEmpty()) {
            contentType()->setMimeType("text/plain");
            contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
        } else if (!htmlBody.isEmpty()) {
            contentType()->setMimeType("text/html");
            contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
        } else if (dynamic_cast<MapiEmbeddedNote*>(this)) {
            contentType()->setMimeType("message/rfc822");
            contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
        } else {
            // Give up.
        }
    }

    KMime::Content *parent = this;
    KMime::Content *body;
    if (!textBody.isEmpty() && !htmlBody.isEmpty()) {
        if (parent->contentType()->mimeType() != "multipart/alternative") {
            body = new KMime::Content;
            body->contentType()->setMimeType("multipart/alternative");
            body->contentType()->setBoundary(KMime::multiPartBoundary());
            parent->addContent(body);
            parent = body;
        }

        body = new KMime::Content;
        body->contentType()->setMimeType("text/plain");
        body->contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
        body->setBody(textBody.toUtf8());
        parent->addContent(body);

        body = new KMime::Content;
        body->contentType()->setMimeType("text/html");
        body->contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
        body->setBody(htmlBody.toUtf8());
        parent->addContent(body);
    } else if (!textBody.isEmpty()) {
        if (parent->contentType()->mimeType() == "text/plain") {
            parent->setBody(textBody.toUtf8());
        } else {
            body = new KMime::Content;
            body->contentType()->setMimeType("text/plain");
            body->contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
            body->setBody(textBody.toUtf8());
            parent->addContent(body);
        }
    } else if (!htmlBody.isEmpty()) {
        if (parent->contentType()->mimeType() == "text/html") {
            parent->setBody(htmlBody.toUtf8());
        } else {
            body = new KMime::Content;
            body->contentType()->setMimeType("text/html");
            body->contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
            body->setBody(htmlBody.toUtf8());
            parent->addContent(body);
        }
    } else {
        // No body to speak of...
    }

    // Short circuit exit if there are no attachments.
    if (!hasAttachments) {
        assemble();
        return true;
    }
    if (MAPI_E_SUCCESS != GetAttachmentTable(&m_object, &m_attachments)) {
        error() << "cannot get attachment table:" << mapiError();
        return false;
    }
    // The list of tags used to fetch an attachment, from [MS-OXCMSG].
    static int attachmentTagList[] = {
        // 2.2.2.6
        PidTagAttachNumber,
        // 2.2.2.7
        PidTagAttachDataBinary,
        // 2.2.2.8
        PidTagAttachDataObject,
        // 2.2.2.9
        PidTagAttachMethod,
        // 2.2.2.10
        PidTagAttachLongFilename,
        // 2.2.2.11
        PidTagAttachFilename,
        // 2.2.2.16
        PidTagRenderingPosition,
        // 2.2.2.25
        PidTagTextAttachmentCharset,
        // 2.2.2.26
        PidTagAttachMimeTag,
        PidTagAttachContentId,
        PidTagAttachContentLocation,
        PidTagAttachContentBase,
        0 };
    static SPropTagArray attachmentTags = {
        (sizeof(attachmentTagList) / sizeof(attachmentTagList[0])) - 1,
        (MAPITAGS *)attachmentTagList };

    if (MAPI_E_SUCCESS != SetColumns(&m_attachments, &attachmentTags)) {
        error() << "cannot set attachment table columns:" << mapiError();
        return false;
    }

    // Get current cursor position.
    uint32_t cursor;
    if (MAPI_E_SUCCESS != QueryPosition(&m_attachments, NULL, &cursor)) {
        error() << "cannot query attachments position:" << mapiError();
        return false;
    }

    // Iterate through sets of rows.
    SRowSet rowset;
    while ((QueryRows(&m_attachments, cursor, TBL_ADVANCE, &rowset) == MAPI_E_SUCCESS) && rowset.cRows) {
        for (unsigned i = 0; i < rowset.cRows; i++) {
            SRow &row = rowset.aRow[i];
            unsigned number = 0;
            unsigned renderingPosition = 0;
            QString file;
            unsigned method = 0;
            QString charset;
            QString mimeTag = QString::fromAscii("text/plain");
            QString contentId;
            QString contentLocation;
            QString contentBase;

            for (unsigned j = 0; j < row.cValues; j++) {
                MapiProperty property(row.lpProps[j]); 

                // Note that the set of properties fetched here must be aligned
                // with those set above.
                switch (property.tag()) {
                case PidTagAttachNumber: 
                    number = property.value().toUInt();
                    break;
                case PidTagAttachDataBinary: 
                case PidTagAttachDataObject: 
                    break;
                case PidTagAttachMethod: 
                    method = property.value().toUInt();
                    break;
                case PidTagAttachLongFilename: 
                    file = property.value().toString();
                    break;
                case PidTagAttachFilename:
                    if (file.isEmpty()) {
                        file = property.value().toString();
                    }
                    break;
                case PidTagRenderingPosition: 
                    renderingPosition = property.value().toUInt();
                    break;
                case PidTagTextAttachmentCharset:
                    charset = property.value().toString();
                    break;
                case PidTagAttachMimeTag: 
                    mimeTag = property.value().toString();
                    break;
                case PidTagAttachContentId:
                    contentId = property.value().toString();
                    break;
                case PidTagAttachContentLocation:
                    contentLocation = property.value().toString();
                    break;
                case PidTagAttachContentBase:
                    contentBase = property.value().toString();
                    break;
                default:
#if (DEBUG_NOTE_PROPERTIES)
                    debug() << "ignoring attachment property:" << tagName(property.tag()) << property.toString();
#endif
                    break;
                }
            }

            QByteArray bytes;
            KMime::Content *attachment;
            MapiEmbeddedNote *embeddedMsg;
            switch (method)
            {
            case ATTACH_BY_VALUE:
                if (UINT_MAX > (index = propertyFind(PidTagAttachDataBinary))) {
                    bytes = propertyAt(index).toByteArray();
                } else {
                    if (MAPI_E_SUCCESS != OpenAttach(&m_object, number, &m_attachment)) {
                        error() << "cannot open attachment" << mapiError();
                        return false;
                    }
                    if (!streamRead(&m_attachment, PidTagAttachDataBinary, bytes)) {
                        return false;
                    }
                }
                attachment = new KMime::Content;

                // Write the attachment as per the rules in [MS-OXCMAIL] 2.1.3.4.
                attachment->contentType()->setMimeType(mimeTag.toUtf8());
                attachment->contentTransferEncoding()->setEncoding(KMime::Headers::CEbase64);
                if (!charset.isEmpty()) {
                    attachment->contentType()->setCharset(charset.toUtf8());
                }
                if (!contentId.isEmpty()) {
                    attachment->contentID()->setIdentifier(contentId.toUtf8());
                } else {
                    if (!contentLocation.isEmpty()) {
                        attachment->contentLocation()->fromUnicodeString(contentLocation, "utf-8");
                    }
                    if (!contentBase.isEmpty()) {
                        //attachment->contentBase()->setCharset(contentBase.toUtf8());
                    }
                }
                if (!file.isEmpty()) {
                    attachment->contentDescription()->fromUnicodeString(file, "utf-8");
                }
                attachment->setBody(bytes);
                addContent(attachment);
                break;
            case ATTACH_EMBEDDED_MSG:
                if (MAPI_E_SUCCESS != OpenAttach(&m_object, number, &m_attachment)) {
                    error() << "cannot open embedded attachment" << mapiError();
                    return false;
                }
                {
                MapiId attachmentId(m_id, (mapi_id_t)number);
                embeddedMsg = new MapiEmbeddedNote(m_connection, "MapiEmbeddedNote", attachmentId, &m_attachment);
                }
                if (!embeddedMsg->open()) {
                    return false;
                }
                if (!embeddedMsg->propertiesPull()) {
                    return false;
                }

                // Write the attachment as per the rules in [MS-OXCMAIL] 2.1.3.4.
                // We need to use setBody() to install the content for
                // message/rfc822, and if there is no subsequent addContent(),
                // a parse() + assemble() sequence is needed.
                if (contentType()->mimeType() == "message/rfc822") {
                    setBody(embeddedMsg->encodedContent());
                    parse();
                    assemble();
                } else {
                    body = new KMime::Content;
                    body->contentType()->setMimeType(mimeTag.toUtf8());
                    body->contentTransferEncoding()->setEncoding(KMime::Headers::CE7Bit);
                    body->contentDisposition()->from7BitString("inline");
                    body->setBody(embeddedMsg->encodedContent());
                    addContent(body);
                }
                break;
            default:
                error() << "ignoring attachment method:" << method;
                break;
            }

            // We are going to reuse this object. Make sure we don't leak stuff.
            mapi_object_release(&m_attachment);
            mapi_object_init(&m_attachment);
        }
    }
    assemble();
    return true;
}

bool MapiNote::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
    /**
     * The list of tags used to fetch a Note, based on [MS-OXCMSG].
     */
    static int ourTagList[] = {
        // 2.2.1.2
        //PidTagHasAttachments,
        // 2.2.1.3
        PidTagMessageClass,
        // 2.2.1.4
        PidTagMessageCodepage,
        // 2.2.1.5
    //	PidTagMessageLocaleId,
        // 2.2.1.6
        PidTagMessageFlags,
        // 2.2.1.7
        //PidTagMessageSize,
        // 2.2.1.8
        //PidTagMessageStatus,
        // 2.2.1.9
        //PidTagSubjectPrefix,
        // 2.2.1.10
        //PidTagNormalizedSubject,
        // 2.2.1.11
        //PidTagImportance,
        // 2.2.1.12
        //PidTagPriority,
        // 2.2.1.13
        //PidTagSensitivity,
        // 2.2.1.14
        //PidLidSmartNoAttach,
        // 2.2.1.15
        //PidLidPrivate,
        // 2.2.1.16
        //PidLidSideEffects,
        // 2.2.1.17
        //PidNameKeywords,
        // 2.2.1.18
        //PidLidCommonStart,
        // 2.2.1.19
        //PidLidCommonEnd,
        // 2.2.1.20
        //PidTagAutoForwarded,
        // 2.2.1.21
        //PidTagAutoForwardComment,
        // 2.2.1.22
        //PidLidCategories,
        // 2.2.1.23
        //PidLidClassification,
        // 2.2.1.24
        //PidLidClassificationDescription,
        // 2.2.1.25
        //PidLidClassified,
        // 2.2.1.26
        //PidTagInternetReferences,
        // 2.2.1.27
        //PidLidInfoPathFormName,
        // 2.2.1.28
        //PidTagMimeSkeleton,
        // 2.2.1.29
        //PidTagTnefCorrelationKey,
        // 2.2.1.30
        //PidTagAddressBookDisplayNamePrintable,
        // 2.2.1.31
        //PidTagCreatorEntryId,
        // 2.2.1.32
        //PidTagLastModifierEntryId,
        // 2.2.1.33
        //PidLidAgingDontAgeMe,
        // 2.2.1.34
        //PidLidCurrentVersion,
        // 2.2.1.35
        //PidLidCurrentVersionName,
        // 2.2.1.36
        //PidTagAlternateRecipientAllowed,
        // 2.2.1.37
        //PidTagResponsibility,
        // 2.2.1.38
        //PidTagRowid,
        // 2.2.1.39
        //PidTagHasNamedProperties,
        // 2.2.1.40
        //PidTagRecipientOrder,
        // 2.2.1.41
        //PidNameContentBase,
        // 2.2.1.42
        //PidNameAcceptLanguage,
        // 2.2.1.43
        //PidTagPurportedSenderDomain,
        // 2.2.1.44
        //PidTagStoreEntryId,
        // 2.2.1.45
        //PidTagTrustSender,
        // 2.2.1.46
        //PidTagSubject,
        // 2.2.1.47
        //PidTagMessageRecipients,
        // 2.2.1.48.1
        PidTagBody,
        // 2.2.1.48.2
        //PidTagNativeBody,
        // 2.2.1.48.3
        //PidTagBodyHtml,
        // 2.2.1.48.4
        //PidTagRtfCompressed,
        // 2.2.1.48.5
        //PidTagRtfInSync,
        // 2.2.1.48.6
        //PidTagInternetCodepage,
        // 2.2.1.48.7
        //PidTagBodyContentId,
        // 2.2.1.48.8
        //PidTagBodyContentLocation,
        // 2.2.1.48.9
        PidTagHtml,
        // 2.2.2.3
        //PidTagCreationTime,
        // ???
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

bool MapiNote::propertiesPull()
{
    static bool tagsAppended = false;
    static QVector<int> tags;

    if (!propertiesPull(tags, tagsAppended, (DEBUG_NOTE_PROPERTIES) != 0)) {
        tagsAppended = true;
        return false;
    }
    tagsAppended = true;
    return true;
}

bool MapiNote::propertiesPush()
{
    // Overwrite all the fields we know about.
#if 0
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

AKONADI_RESOURCE_MAIN( ExMailResource )

#include "exmailresource.moc"
