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

#ifndef MAPICONNECTOR2_H
#define MAPICONNECTOR2_H

#include <QBitArray>
#include <QDateTime>
#include <QDebug>
#include <QList>
#include <QMap>
#include <QString>

extern "C" {
// libmapi is a C library and must therefore be included that way
// otherwise we'll get linker errors due to C++ name mangling
#include <libmapi/libmapi.h>
}

/**
 * Enumerate all starting points in the MAPI store.
 */
typedef enum
{
	MailboxRoot             = olFolderMailboxRoot,		// To navigate the whole tree.
	TopInformationStore     = olFolderTopInformationStore,	// All email note items.
	DeletedItems            = olFolderDeletedItems,
	Outbox                  = olFolderOutbox,
	SentMail                = olFolderSentMail,
	Inbox                   = olFolderInbox,
	CommonView              = olFolderCommonView,
	Calendar                = olFolderCalendar,		// All calendar items.
	Contacts                = olFolderContacts,
	Journal                 = olFolderJournal,
	Notes                   = olFolderNotes,
	Tasks                   = olFolderTasks,
	Drafts                  = olFolderDrafts,
	FoldersAllPublicFolders = olPublicFoldersAllPublicFolders,
	Conflicts               = olFolderConflicts,
	SyncIssues              = olFolderSyncIssues,
	LocalFailures           = olFolderLocalFailures,
	ServerFailures          = olFolderServerFailures,
	Junk                    = olFolderJunk,
	Finder                  = olFolderFinder,
	PublicRoot              = olFolderPublicRoot,
	PublicIPMSubtree        = olFolderPublicIPMSubtree,
	PublicNonIPMSubtree     = olFolderPublicNonIPMSubtree,
	PublicEFormsRoot        = olFolderPublicEFormsRoot,
	PublicFreeBusyRoot      = olFolderPublicFreeBusyRoot,
	PublicOfflineAB         = olFolderPublicOfflineAB,
	PublicEFormsRegistry    = olFolderPublicEFormsRegistry,
	PublicLocalFreeBusy     = olFolderPublicLocalFreeBusy,
	PublicLocalOfflineAB    = olFolderPublicLocalOfflineAB,
	PublicNNTPArticle       = olFolderPublicNNTPArticle,
} MapiDefaultFolder;

/**
 * The id for a MAPI object is (a) hierarchical and (b) associated with a 
 * provider. The provider is set with reference to the @ref MapiDefaultFolder
 * and then cascaded down to contained items as we recurse down a folder tree. 
 */
class MapiId : public QPair<mapi_id_t, mapi_id_t>
{
public:
	/**
	 * For a default folder.
	 */
	MapiId(class MapiConnector2 *connection, MapiDefaultFolder folderType);

	/**
	 * For a child object.
	 */
	MapiId(const MapiId &parent, const mapi_id_t &child);

	/**
	 * From an Akonadi id.
	 */
	MapiId(const QString &id);

	/**
	 * To Akonadi id.
	 */
	QString toString() const;

	bool isValid() const;

	static const QChar fidIdSeparator;

private:
	enum Provider
	{
		INVALID,
		EMSDB,
		NSPI
	} m_provider;
	friend class MapiConnector2;
};

/**
 * Stringified errors.
 */
extern QString mapiError();

/**
 * Try to extract an email address from a string.
 * 
 * @param source	The source to extract from.
 * @param type		The type of source. Conceptually, this is based on the 
 * 			documented values of PidTagAddressType (3COM, ATT,
 * 			CCMAIL, COMPUSERVE, EX, FAX, MSFAX, MCI, MHS, MS, MSA, 
 * 			MSN, PROFS, SMTP, SNADS, TELEX, X400, X500) though we
 * 			only actually support some of these.
 * @param emptyDefault	If no email address can be extracted, set the result to
 * 			an empty string is true, or the original source.
 * @return An email address if we can, or a default as per @ref emptyDefault.
 */
extern QString mapiExtractEmail(const QString &source, const QByteArray &type, bool emptyDefault = false);

extern QString mapiExtractEmail(const class MapiProperty &source, const QByteArray &type, bool emptyDefault = false);

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

class MapiRecipient
{
public:
	typedef enum {
		Sender = 0,
		To = 1,
		CC = 2,
		BCC = 3,
		ReplyTo = 4
	} Type;

	/**
	 * Legal values for the display type. Other values are not valid,
	 * but will be stringified by @ref displayTypeString.
	 */
	typedef enum {
		DtMailuser =		DT_MAILUSER,
		DtDistlist =		DT_DISTLIST,
		DtForum =		DT_FORUM,
		DtAgent =		DT_AGENT,
		DtOrganization =	DT_ORGANIZATION,
		DtPrivateDistlist =	DT_PRIVATE_DISTLIST,
		DtRemoteMailuser =	DT_REMOTE_MAILUSER,
		DtRoom =		0x7,
		DtEquipment =		0x8,
		DtSecurityGroup =	0x9,
	} DisplayType;

	/**
	 * An i18n value for the type.
	 */
	static QString toString(DisplayType type, unsigned pluralForm = 1);

	/**
	 * Legal values for the type of object. Other values are not valid,
	 * but will be stringified by @ref objectTypeString.
	 */
	typedef enum {
		OtStore =		MAPI_STORE,
		OtAddrbook =		MAPI_ADDRBOOK,
		OtFolder =		MAPI_FOLDER,
		OtABcont =		MAPI_ABCONT,
		OtMessage =		MAPI_MESSAGE,
		OtMailuser =		MAPI_MAILUSER,
		OtAttach =		MAPI_ATTACH,
		OtDistlist =		MAPI_DISTLIST,
		OtProfsect =		MAPI_PROFSECT,
		OtStatus =		MAPI_STATUS,
		OtSession =		MAPI_SESSION,
		OtForminfo =		MAPI_FORMINFO,
	} ObjectType;

	MapiRecipient(Type type) :
		m_type(type),
		m_displayType(DtMailuser),
		m_objectType(OtMailuser)
	{
		trackStatus = 0;
		flags = 0;
		order = 0;
	}

	void setType(Type type)
	{
		m_type = type;
	}

	Type type() const
	{
		return m_type;
	}

	/**
	 * Map all recipient types to strings.
	 */
	QString typeString() const;

	void setObjectType(ObjectType type)
	{
		m_objectType = type;
	}

	ObjectType objectType() const
	{
		return m_objectType;
	}

	/**
	 * Map all MAPI object types to strings.
	 */
	QString objectTypeString() const;

	void setDisplayType(DisplayType type)
	{
		m_displayType = type;
	}

	DisplayType displayType() const
	{
		return m_displayType;
	}

	/**
	 * Map all MAPI display types to strings.
	 */
	QString displayTypeString() const;

	QString name;
	QString email;
	unsigned trackStatus;
	unsigned flags;
	unsigned order;

	/**
	 * Convert the whole things to a string for debug purposes.
	 */
	QString toString() const;

protected:
	Type m_type;
	DisplayType m_displayType;
	ObjectType m_objectType;
};

class MapiRecurrencyPattern 
{
public:
	enum RecurrencyType {
		Daily, Weekly, Every_Weekday, Monthly, Yearly,
	};
	enum EndTyp {
		Never, Count, Date,
	};

	MapiRecurrencyPattern() { mRecurring= false; }
	virtual ~MapiRecurrencyPattern() {}

	bool isRecurring() { return mRecurring; }
	void setRecurring(bool r) { this->mRecurring = r; }

	bool setData(RecurrencePattern* pattern);
	
private:
	QBitArray getRecurrenceDays(const uint32_t exchangeDays);
	int convertDayOfWeek(const uint32_t exchangeDayOfWeek);
	QDateTime convertExchangeTimes(const uint32_t exchangeMinutes);

	bool mRecurring;

public:
	RecurrencyType mRecurrencyType;
	int mPeriod;
	int mFirstDOW;
	EndTyp mEndType;
	int mOccurrenceCount;
	QBitArray mDays;
	QDateTime mStartDate;
	QDateTime mEndDate;
};

/**
 * A class which wraps a talloc memory allocator such that objects of this type
 * automatically free the used memory on destruction.
 */
class TallocContext
{
public:
	TallocContext(const char *name);

	virtual ~TallocContext();

	TALLOC_CTX *ctx();

	/**
	 * Make a talloc-friendly allocation.
	 */
	template <class T>
	T *allocate();

	/**
	 * Make a talloc-friendly array.
	 */
	template <class T>
	T *array(unsigned size);

	/**
	 * Make a talloc-friendly UTF-8 encoded string.
	 */
	char *string(const QString &original);

protected:
	TALLOC_CTX *m_ctx;

	/**
	 * Debug and error reporting. Each subclass should reimplement with 
	 * logic that emits a prefix identifying the object involved. 
	 */
	virtual QDebug debug() const = 0;
	virtual QDebug error() const = 0;

	/**
	 * A couple of helper functions that subclasses can use to implement the
	 * above virtual methods.
	 */
	QDebug debug(const QString &caller) const;
	QDebug error(const QString &caller) const;
};

/**
 * A class for managing access to MAPI profiles. This is the root for all other
 * MAPI interactions.
 */
class MapiProfiles : protected TallocContext
{
public:
	MapiProfiles();
	virtual ~MapiProfiles();

	/**
	 * Find existing profiles.
	 */
	QStringList list();

	/**
	 * Set the default profile.
	 */
	bool defaultSet(QString profile);

	/**
	 * Get the default profile.
	 */
	QString defaultGet();

	/**
	 * Add a new profile.
	 */
	bool add(QString profile, QString username, QString password, QString domain, QString server);

	/**
	 * Remove a profile.
	 */
	bool remove(QString profile);

protected:
	mapi_context *m_context;

	/**
	 * Must be called first!
	 */
	bool init();

private:
	bool m_initialised;

	bool addAttribute(const char *profile, const char *attribute, QString value);

	virtual QDebug debug() const;
	virtual QDebug error() const;
};

/**
 * The main class represents a connection to the MAPI server.
 */
class MapiConnector2 : public MapiProfiles
{
public:
	MapiConnector2();
	virtual ~MapiConnector2();

	/**
	 * Connect to the server.
	 */
	bool login(QString profile);

	/**
	 * Factory for getting default folder ids. MAPI has two kinds of folder,
	 * Public and user-specific. This wraps the two together.
	 */
	bool defaultFolder(MapiDefaultFolder folderType, MapiId *id);

	/**
	 * How many entries are there in the GAL?
	 */
	bool GALCount(unsigned *totalCount);

	/**
	 * Fetch upto the requested number of entries from the GAL. The start
	 * point is where we previously left off.
	 */
	bool GALRead(unsigned requestedCount, SPropTagArray *tags, SRowSet **results, unsigned *percentagePosition = 0);

	bool GALSeek(const QString &displayName, unsigned *percentagePosition = 0, SPropTagArray *tags = 0, SRowSet **results = 0);

	bool GALRewind();

	mapi_object_t *store(const MapiId &id)
	{
		switch (id.m_provider)
		{
		case MapiId::EMSDB:
			return &m_store;
		case MapiId::NSPI:
			return &m_nspiStore;
		default:
			return 0;
		}
	}

	/**
	 * Resolve the given names.
	 * 
	 * @param names A 0-terminated array of 0-terminated names to be resolved.
	 * @param tags  The item properties we are interested in.
	 */
	bool resolveNames(const char *names[], SPropTagArray *tags,
			  SRowSet **results, PropertyTagArray_r **statuses);

private:
	mapi_object_t openFolder(mapi_id_t folderID);

	mapi_session *m_session;
	mapi_object_t m_store;
	mapi_object_t m_nspiStore;

	virtual QDebug debug() const;
	virtual QDebug error() const;
};

/**
 * A class which wraps a MAPI object such that objects of this type 
 * automatically free the used memory on destruction.
 */
class MapiObject : protected TallocContext
{
public:
	MapiObject(MapiConnector2 *connection, const char *tallocName, const MapiId &id);

	virtual ~MapiObject();

	mapi_object_t *d() const;

	const MapiId &id() const;

	virtual bool open() = 0;

	/**
	 * Add a property with the given int.
	 */
	bool propertyWrite(int tag, int data, bool idempotent = true);

	/**
	 * Add a property with the given string.
	 */
	bool propertyWrite(int tag, QString &data, bool idempotent = true);

	/**
	 * Add a property with the given datetime.
	 */
	bool propertyWrite(int tag, QDateTime &data, bool idempotent = true);

	/**
	 * Set the written properties onto the object, and prepare to go again.
	 */
	virtual bool propertiesPush();

	/**
	 * How many properties do we have?
	 */
	unsigned propertyCount() const;

	/**
	 * Find a property by tag.
	 * 
	 * @return The index, or UINT_MAX if not found.
	 */
	unsigned propertyFind(int tag) const;

	/**
	 * Fetch a property by tag.
	 */
	QVariant property(int tag) const;

	/**
	 * Fetch a property by index.
	 */
	QVariant propertyAt(unsigned i) const;

	/**
	 * Fetch a tag by index.
	 */
	QString tagAt(unsigned i) const;

	/**
	 * For display purposes, convert a property into a string, taking
	 * care to hex-ify GUIDs and other byte arrays, and lists of the
	 * same.
	 */
	QString propertyString(unsigned i) const;

	/**
	 * Find the name for a tag. If it not a well known one, try a lookup.
	 * Technically, this should only be needed if bit 31 is set, but
	 * still...
	 */
	QString tagName(int tag) const;

protected:
	MapiConnector2 *m_connection;
	const MapiId m_id;
	struct SPropValue *m_properties;
	uint32_t m_propertyCount;
	mutable mapi_object_t m_object;

	/**
	 * Fetch a set of properties.
	 * 
	 * @return Whether the pull succeeds, irrespective of whether the tags
	 * were matched.
	 */
	virtual bool propertiesPull(QVector<int> &tags, bool tagsAppended, bool pullAll);

private:
	/**
	 * Fetch all properties.
	 */
	virtual bool propertiesPull();

	/**
	 * Add a property with the given value, using an immediate assignment.
	 */
	bool propertyWrite(int tag, void *data, bool idempotent = true);

	int *m_ourTagList;
	SPropTagArray m_ourTags;
};

/**
 * Represents a MAPI item. Objects of this type contain enough information to
 * allow the corresponding full item to be retrieved.
 * 
 * @ref MapiFolder
 */
class MapiItem
{
public:
	MapiItem(const MapiId &id, QString &name, QDateTime &modified);

	/**
	 * The id of the full item.
	 */
	const MapiId &id() const;

	/**
	 * The name of this item.
	 */
	QString name() const;

	/**
	 * The last-modified date time of the full item.
	 */
	QDateTime modified() const;

private:
	const MapiId m_id;
	const QString m_name;
	const QDateTime m_modified;
};

/**
 * Represents a MAPI folder. A folder contains other child folder and
 * @ref MapiItem objects.
 */
class MapiFolder : public MapiObject
{
public:
	MapiFolder(MapiConnector2 *connection, const char *tallocName, const MapiId &id);

	virtual ~MapiFolder();

	virtual bool open();

	QString name;

	/**
	 * Fetch children which are folders.
	 * 
	 * @param children	The children will be added to this list. The 
	 * 			caller is responsible for freeing entries on 
	 * 			the list.
	 * @param filter	Only return items whose PR_CONTAINER_CLASS 
	 * 			starts with this value, or the empty string to
	 * 			get all of them.
	 */
	bool childrenPull(QList<MapiFolder *> &children, const QString &filter = QString());

	/**
	 * Fetch children which are not folders.
	 * 
	 * @param children	The children will be added to this list. The 
	 * 			caller is responsible for freeing entries on 
	 * 			the list.
	 */
	bool childrenPull(QList<MapiItem *> &children);

protected:
	mapi_object_t m_contents;

private:
	virtual QDebug debug() const;
	virtual QDebug error() const;
};

/**
 * A Message, with recipients.
 */
class MapiMessage : public MapiObject
{
public:
	MapiMessage(MapiConnector2 *connection, const char *tallocName, const MapiId &id);

	virtual bool open();

	/**
	 * Fetch all properties.
	 */
	virtual bool propertiesPull();

	/**
	 * Lists of To, CC and BCC, as well as the sender (the last should have 
	 * 0 or 1 items only, but in theory may have more).
	 */
	const QList<MapiRecipient> &recipients();

	/**
	 * Add another recipient to our list.
	 * 
	 * @param source 	A descriptive string identifying the source, for
	 * 			debug purposes.
	 * @param candidate	A recipient. If unique, it will be added. If not,
	 * 			the current entry will be updated where appropriate
	 * 			with information from the candidate.
	 */
	void addUniqueRecipient(const char *source, MapiRecipient &candidate);

protected:
	QList<MapiRecipient> m_recipients;

	/**
	 * Pull a given set of properties, plus any we need internally.
	 * 
	 * @param tags		Properties to pull.
	 * @param tagsAppended	False if the set of tags given have not already
	 * 			been augmented with the one we need need.
	 * 			Used by the caller to turn the construction of 
	 * 			the set of tags into a one-time operation.
	 * @param pullAll	If true, all available properties will be pulled
	 * 			rather than just the given @ref tags.
	 */
	virtual bool propertiesPull(QVector<int> &tags, const bool tagsAppended, bool pullAll);

private:
	virtual QDebug debug() const;
	virtual QDebug error() const;

	/**
	 * Fetch all recipients.
	 */
	bool recipientsPull();

	/**
	 * Flesh out a recipient.
	 */
	void recipientPopulate(const char *phase, SRow &recipient, MapiRecipient &result);
};

#endif // MAPICONNECTOR2_H
