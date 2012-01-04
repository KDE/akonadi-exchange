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
#include <kpimutils/email.h>

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
#ifndef DEBUG_NOTE_PROPERTIES
#define DEBUG_NOTE_PROPERTIES 0
#endif

using namespace Akonadi;

/**
 * An Email.
 */
class MapiNote : public MapiMessage, public KMime::Message
{
public:
	MapiNote(MapiConnector2 *connection, const char *tallocName, mapi_id_t folderId, mapi_id_t id);

	virtual ~MapiNote();

	/**
	 * Fetch all note properties.
	 */
	virtual bool propertiesPull();

	/**
	 * Update a note item.
	 */
	virtual bool propertiesPush();

private:
	virtual QDebug debug() const;
	virtual QDebug error() const;

	bool preparePayload();

	/**
	 * Fetch email properties.
	 */
	virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);

	/**
	 * Dump a change to a header.
	 */
	void dumpChange(KMime::Headers::Base *header, const char *item, MapiProperty &property);
	void dumpChange(KMime::Headers::Base *header, const char *item, MapiProperty &property, QString &value);

	mapi_object_t m_attachments;
	mapi_object_t m_attachment;
	mapi_object_t m_stream;
};

/**
 * The list of tags used to fetch an attachment.
 */
static int attachmentTagList[] = {
	PidTagAttachNumber,
	PidTagRenderingPosition,
	PidTagAttachMimeTag,
	PidTagAttachMethod,
	PidTagAttachLongFilename,
	PidTagAttachFilename,
	PidTagAttachSize,
	PidTagAttachDataBinary,
	PidTagAttachDataObject,
	0 };
static SPropTagArray attachmentTags = {
	(sizeof(attachmentTagList) / sizeof(attachmentTagList[0])) - 1,
	(MAPITAGS *)attachmentTagList };

	
/**
 * Take a set of properties, and attempt to apply them to the given addressee.
 * 
 * The switch statement at the heart of this routine must be kept synchronised
 * with @ref noteTagList.
 *
 * @return false on error.
 */
void MapiNote::dumpChange(KMime::Headers::Base *header, const char *item, MapiProperty &property)
{
	if (header) {
		error() << "change" << item << header->asUnicodeString() << "using" << tagName(property.tag()) << property.value().toString();
	}
}
void MapiNote::dumpChange(KMime::Headers::Base *header, const char *item, MapiProperty &property, QString &value)
{
	if (header) {
		error() << "change" << item << header->asUnicodeString() << "using" << tagName(property.tag()) << value;
	}
}

bool MapiNote::preparePayload()
{
/*
	QString title;
	QString text;
	QString sender;
	QDateTime created;

	
	QStringList displayTo;
*/

	bool hasAttachments = false;
	unsigned index;
	KMime::Content *body;

	// First set the header content, and parse what we can from it. Note
	// this anything beyond the header will return the last value only. For
	// example, this will result in a contentType of "text/html" instead of
	// "multipart/alternative":
	//
	// Content-Type: multipart/alternative;^M
	// boundary="----_=_NextPart_001_01CC9A5A.689896D8"^M
	// Subject: ...^M
	// ...
	// Return-Path: owner-build-guru@mtv-core-1.cisco.com^M
	// ^M
	// ------_=_NextPart_001_01CC9A5A.689896D8^M
	// Content-Type: text/plain;^M
	//         charset="US-ASCII"^M
	// Content-Transfer-Encoding: quoted-printable^M
	// ^M
	// ------_=_NextPart_001_01CC9A5A.689896D8^M
	// Content-Type: text/html;^M
	//         charset="US-ASCII"^M
	if (UINT_MAX > (index = propertyFind(PidTagTransportMessageHeaders))) {
		//setHead(propertyAt(index).value().toString().toUtf8());
		//parse();
	}
	kError() << "+++++++++ before" << contentType()->asUnicodeString();
	// Walk through the properties and extract the values of interest. The
	// properties here should be aligned with the list pulled above.
	for (unsigned i = 0; i < m_propertyCount; i++) {
		MapiProperty property(m_properties[i]);
		QString email;
		QString name;

		switch (property.tag()) {
		case PidTagMessageClass:
			// Sanity check the message class.
			if ((QLatin1String("IPM.Note") != property.value().toString()) &&
				(QLatin1String("Remote.IPM.Note") != property.value().toString())){
				error() << "retrieved item is not an email or a header:" << property.value().toString();
				return false;
			}
			break;
		// 2.2.2.3
		case PidTagCreationTime:
			dumpChange(date(true), "date", property);
			date()->setDateTime(KDateTime(property.value().toDateTime()));
			break;
		case PidTagTransportMessageHeaders:
			kError() << "headers:" << property.value().toString();
			setHead(property.value().toString().toUtf8());
	kError() << "+++++++++ before parse" << contentType()->asUnicodeString();
	//parse();
	kError() << "+++++++++ after parse" << contentType()->asUnicodeString();
			break;
		case PidTagBody:
			kError() << "text body:" << property.value().toString();
			body = new KMime::Content;
			body->contentType()->setMimeType("text/plain");
			body->setBody(property.value().toString().toUtf8());
			addContent(body);
			break;
		case PidTagHtml:
			kError() << "html body:" << property.value().toString();
			body = new KMime::Content;
			body->contentType()->setMimeType("text/html");
			body->setBody(property.value().toString().toUtf8());
			addContent(body);
			break;
		case PidTagHasAttachments:
			hasAttachments = true;
			break;

		case PidTagConversationTopic:
			if (!subject() || subject()->isEmpty()) {
				dumpChange(subject(true), "subject", property);
				subject()->fromUnicodeString(property.value().toString(), "utf-8");
			}
			break;
		case PidTagSubject:
			dumpChange(subject(true), "subject", property);
			subject()->fromUnicodeString(property.value().toString(), "utf-8");
			break;

			
		case PidTagInternetMessageId:
			//dumpChange(messageID(true), "messageID", property);
			messageID()->fromUnicodeString(property.value().toString(), "utf-8");
			break;
		case PidTagInternetReferences:
			//dumpChange(references(true), "references", property);
			references()->fromUnicodeString(property.value().toString(), "utf-8");
			break;

			
/*
		case PidTagSubject:
			break;
			
PidTagCreationTime (section 2.2.2.3)

PidTagLastModificationTime (section 2.2.2.2)

PidTagLastModifierName ([MS-OXCPRPT] section 2.2.1.5)

PidTagObjectType ([MS-OXCPRPT] section 2.2.1.7)




2.2.1.2 PidTagHasAttachments Property


2.2.1.4 PidTagMessageCodepage Property

2.2.1.5 PidTagMessageLocaleId Property

2.2.1.6 PidTagMessageFlags Property

2.2.1.7 PidTagMessageSize Property

2.2.1.8 PidTagMessageStatus Property

2.2.1.9 PidTagSubjectPrefix Property

2.2.1.10 PidTagNormalizedSubject Property

2.2.1.11 PidTagImportance Property

2.2.1.12 PidTagPriority Property

2.2.1.13 PidTagSensitivity Property

2.2.1.14 PidLidSmartNoAttach Property

2.2.1.15 PidLidPrivate Property

2.2.1.16 PidLidSideEffects Property

2.2.1.17 PidNameKeywords Property

2.2.1.18 PidLidCommonStart Property

2.2.1.19 PidLidCommonEnd Property

2.2.1.20 PidTagAutoForwarded Property

2.2.1.21 PidTagAutoForwardComment Property

2.2.1.22 PidLidCategories Property

2.2.1.23 PidLidClassification

2.2.1.24 PidLidClassificationDescription Property

2.2.1.25 PidLidClassified Property

2.2.1.26 PidTagInternetReferences Property

2.2.1.27 PidLidInfoPathFormName Property

2.2.1.28 PidTagMimeSkeleton Property

2.2.1.29 PidTagTnefCorrelationKey Property

2.2.1.30 PidTagAddressBookDisplayNamePrintable Property

2.2.1.31 PidTagCreatorEntryId Property

2.2.1.32 PidTagLastModifierEntryId Property

2.2.1.33 PidLidAgingDontAgeMe Property

2.2.1.34 PidLidCurrentVersion Property

2.2.1.35 PidLidCurrentVersionName Property

2.2.1.36 PidTagAlternateMapiRecipientAllowed Property

2.2.1.37 PidTagResponsibility Property

2.2.1.38 PidTagRowid Property

2.2.1.39 PidTagHasNamedProperties Property

2.2.1.40 PidTagMapiRecipientOrder Property

2.2.1.41 PidNameContentBase Property

2.2.1.42 PidNameAcceptLanguage Property

2.2.1.43 PidTagPurportedSenderDomain Property

2.2.1.44 PidTagStoreEntryId Property

2.2.1.45 PidTagTrustSender

2.2.1.46  Property

2.2.1.47 PidTagMessageMapiRecipients Property

2.2.1.48 Body Properties

2.2.1.49 Contact Linking Properties

2.2.1.50 Retention and Archive Properties

*/



/*
			text = property.value().toString();
			break;
		case PidTagCreationTime:
			created = property.value().toDateTime();
			break;
*/
		default:
#if (DEBUG_NOTE_PROPERTIES)
			debug() << "ignoring note property:" << tagName(property.tag()) << property.toString();
#endif
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

	// Short circuit exit if there are no attachments.
	if (!hasAttachments) {
		assemble();
		return true;
	}
	if (MAPI_E_SUCCESS != GetAttachmentTable(&m_object, &m_attachments)) {
		error() << "cannot get attachment table:" << mapiError();
		return false;
	}
	error() << "eErr got attachment table";
	if (MAPI_E_SUCCESS != SetColumns(&m_attachments, &attachmentTags)) {
		error() << "cannot set attachment table columns:" << mapiError();
		return false;
	}
	error() << "eErr set cols table";

	while (true) {
		// Get current cursor position.
		uint32_t cursor;
		if (MAPI_E_SUCCESS != QueryPosition(&m_attachments, NULL, &cursor)) {
			error() << "cannot query attachments position:" << mapiError();
			return false;
		}
		// Iterate through sets of rows.
		SRowSet rowset;

		error() << "get attachments from row:" << cursor;
		if (MAPI_E_SUCCESS != QueryRows(&m_attachments, cursor, TBL_ADVANCE, &rowset)) {
			error() << "cannot query attachments" << mapiError();
			return false;
		}
		if (!rowset.cRows) {
			error() << "attachment count" << cursor;
			break;
		}
		error() << "got rows" << rowset.cRows;
		for (unsigned i = 0; i < rowset.cRows; i++) {
			SRow &row = rowset.aRow[i];
			unsigned number = 0;
			unsigned renderingPosition = 0;
			QString mimeTag;
			QString file;
			unsigned method = 0;
			unsigned dataSize;

			for (unsigned j = 0; j < row.cValues; j++) {
				MapiProperty property(row.lpProps[j]); 

				// Note that the set of properties fetched here must be aligned
				// with those set above.
				switch (property.tag()) {
				case PidTagAttachNumber: 
					number = property.value().toUInt();
					break;
				case PidTagRenderingPosition: 
					renderingPosition = property.value().toUInt();
					break;
				case PidTagAttachMimeTag: 
					mimeTag = property.value().toString();
					break;
				case PidTagAttachLongFilename: 
					file = property.value().toString();
					break;
				case PidTagAttachFilename:
					if (file.isEmpty()) {
						file = property.value().toString();
					}
					break;
				case PidTagAttachMethod: 
					method = property.value().toUInt();
					break;
				case PidTagAttachSize: 
					dataSize = property.value().toUInt();
					error() << "read data size" << dataSize;
					break;
				case PidTagAttachDataBinary: 
					break;
				case PidTagAttachDataObject: 
					break;
				default:
					debug() << "ignoring attachment property:" << tagName(property.tag()) << property.toString();
					break;
				}
			}

			QByteArray attachment;
			unsigned offset;
			uint16_t readSize;
			debug() << "attachment method:" << method;
			switch (method)
			{
			case 1:
				if (MAPI_E_SUCCESS != OpenAttach(&m_object, number, &m_attachment)) {
					error() << "cannot open attachment" << mapiError();
					return false;
				}
				if (MAPI_E_SUCCESS != OpenStream(&m_attachment, (MAPITAGS)PidTagAttachDataBinary, OpenStream_ReadOnly, &m_stream)) {
					error() << "cannot open stream" << mapiError();
					return false;
				}
				if (MAPI_E_SUCCESS != GetStreamSize(&m_stream, &dataSize)) {
					error() << "cannot get stream size" << mapiError();
					return false;
				}
				error() << "fetched stream size" << dataSize;
				attachment.reserve(dataSize);
				offset = 0;
				do {
					if (MAPI_E_SUCCESS != ReadStream(&m_stream, (uchar *)attachment.data() + offset, 1024, &readSize)) {
						error() << "cannot read stream" << mapiError();
						return false;
					}
					offset += readSize;
				} while (readSize && (offset < dataSize));
				attachment.resize(dataSize);
				error() << "stream attachment size:" << attachment.size();
				body = new KMime::Content;
				body->contentType()->setMimeType(mimeTag.toUtf8());
				if (!file.isEmpty()) {
					body->contentDescription(true)->fromUnicodeString(file, "utf-8");
				}
				body->setBody(attachment);
				addContent(body);
				break;
			case 5:
				error() << "PidTagAttachDataBinary" << propertyFind(PidTagAttachDataBinary);
				error() << "PidTagAttachDataObject" << propertyFind(PidTagAttachDataObject);
				error() << "PidTagAttachDataBinary_Error" << propertyFind(PidTagAttachDataBinary_Error);
				error() << "PidTagAttachDataObject_Error" << propertyFind(PidTagAttachDataObject_Error);
				
				if (UINT_MAX > (index = propertyFind(PidTagAttachDataBinary))) {
					attachment = propertyAt(index).toByteArray();
				error() << mimeTag << "attachment object:" << attachment.toHex();
				} else {
					if (MAPI_E_SUCCESS != OpenAttach(&m_object, number, &m_attachment)) {
						error() << "cannot open attachment" << mapiError();
						return false;
					}
					if (MAPI_E_SUCCESS != OpenStream(&m_attachment, (MAPITAGS)PidTagAttachDataObject, OpenStream_ReadOnly, &m_stream)) {
						error() << "cannot open stream" << mapiError();
						return false;
					}
					if (MAPI_E_SUCCESS != GetStreamSize(&m_stream, &dataSize)) {
						error() << "cannot get stream size" << mapiError();
						return false;
					}
					attachment.reserve(dataSize);
					offset = 0;
					do {
						if (MAPI_E_SUCCESS != ReadStream(&m_stream, (uchar *)attachment.data() + offset, 1024, &readSize)) {
							error() << "cannot read stream" << mapiError();
							return false;
						}
						offset += readSize;
					} while (readSize && (offset < dataSize));
					attachment.resize(dataSize);
					error() << mimeTag << "binary attachment size:" << attachment.size();
				}
				kError() << mimeTag << "attachment object:" << attachment.toHex();
				body = new KMime::Content;
				body->contentType()->setMimeType(mimeTag.toUtf8());
				if (!file.isEmpty()) {
					body->contentDescription(true)->fromUnicodeString(file, "utf-8");
				}
				body->setBody(attachment);
				addContent(body);
				break;
			default:
				error() << "ignoring attachment method:" << method;
				break;
			}
		}
	}
	assemble();
	return true;
}

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
#if (DEBUG_NOTE_PROPERTIES)
	while (items.size() > 3) {
		items.removeLast();
	}
#endif
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
//	item.setModificationTime(message->modified);
*/
	item.setPayload<KMime::Message::Ptr>(ptr);

	// Not needed!
	//delete message;

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

MapiNote::MapiNote(MapiConnector2 *connector, const char *tallocName, mapi_id_t folderId, mapi_id_t id) :
	MapiMessage(connector, tallocName, folderId, id)
{
	mapi_object_init(&m_attachments);
	mapi_object_init(&m_attachment);
	mapi_object_init(&m_stream);
}

MapiNote::~MapiNote()
{
	mapi_object_release(&m_stream);
	mapi_object_release(&m_attachment);
	mapi_object_release(&m_attachments);
}

QDebug MapiNote::debug() const
{
	static QString prefix = QString::fromAscii("MapiNote: %1/%2:");
	return MapiObject::debug(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE)) /*<< title*/;
}

QDebug MapiNote::error() const
{
	static QString prefix = QString::fromAscii("MapiNote: %1/%2");
	return MapiObject::error(prefix.arg(m_folderId, 0, ID_BASE).arg(m_id, 0, ID_BASE)) /*<< title*/;
}

bool MapiNote::propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll)
{
	/**
	 * The list of tags used to fetch a Note, based on [MS-OXCMSG].
	 */
	static int ourTagList[] = {
		// 2.2.1.2  
		PidTagHasAttachments,
		// 2.2.1.3  
		PidTagMessageClass,
		// 2.2.1.4  
		PidTagMessageCodepage,
		// 2.2.1.5  
		PidTagMessageLocaleId,
		// 2.2.1.6  
		PidTagMessageFlags,
		// 2.2.1.7  
		PidTagMessageSize,
		// 2.2.1.8  
		PidTagMessageStatus,
		// 2.2.1.9  
		//PidTagSubjectPrefix,
		// 2.2.1.10 
		//PidTagNormalizedSubject,
		// 2.2.1.11 
		PidTagImportance,
		// 2.2.1.12 
		PidTagPriority,
		// 2.2.1.13 
		PidTagSensitivity,
		// 2.2.1.14 
		PidLidSmartNoAttach,
		// 2.2.1.15 
		PidLidPrivate,
		// 2.2.1.16 
		PidLidSideEffects,
		// 2.2.1.17 
		PidNameKeywords,
		// 2.2.1.18 
		PidLidCommonStart,
		// 2.2.1.19 
		PidLidCommonEnd,
		// 2.2.1.20 
		PidTagAutoForwarded,
		// 2.2.1.21 
		PidTagAutoForwardComment,
		// 2.2.1.22 
		PidLidCategories,
		// 2.2.1.23 
		PidLidClassification,
		// 2.2.1.24 
		PidLidClassificationDescription,
		// 2.2.1.25 
		PidLidClassified,
		// 2.2.1.26 
		PidTagInternetReferences,
		// 2.2.1.27 
		PidLidInfoPathFormName,
		// 2.2.1.28 
		PidTagMimeSkeleton,
		// 2.2.1.29 
		PidTagTnefCorrelationKey,
		// 2.2.1.30 
		PidTagAddressBookDisplayNamePrintable,
		// 2.2.1.31 
		PidTagCreatorEntryId,
		// 2.2.1.32 
		PidTagLastModifierEntryId,
		// 2.2.1.33 
		PidLidAgingDontAgeMe,
		// 2.2.1.34 
		PidLidCurrentVersion,
		// 2.2.1.35 
		PidLidCurrentVersionName,
		// 2.2.1.36 
		PidTagAlternateRecipientAllowed,
		// 2.2.1.37 
		PidTagResponsibility,
		// 2.2.1.38 
		PidTagRowid,
		// 2.2.1.39 
		PidTagHasNamedProperties,
		// 2.2.1.40 
		PidTagRecipientOrder,
		// 2.2.1.41 
		PidNameContentBase,
		// 2.2.1.42 
		PidNameAcceptLanguage,
		// 2.2.1.43 
		PidTagPurportedSenderDomain,
		// 2.2.1.44 
		PidTagStoreEntryId,
		// 2.2.1.45 
		PidTagTrustSender,
		// 2.2.1.46 
		PidTagSubject,
		// 2.2.1.47 
		PidTagMessageRecipients,
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
		PidTagInternetCodepage,
		// 2.2.1.48.7
		PidTagBodyContentId,
		// 2.2.1.48.8
		PidTagBodyContentLocation,
		// 2.2.1.48.9
		PidTagHtml,
		// 2.2.2.3 
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
