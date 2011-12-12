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

class Recipient
{
public:
	Recipient()
	{
		trackStatus = 0;
		flags = 0;
		type = 0;
		order = 0;
	}

	QString name;
	QString email;
	unsigned trackStatus;
	unsigned flags;
	unsigned type;
	unsigned order;
};

class Attendee : public Recipient
{
public:
	Attendee() :
		Recipient()
	{
	}

	Attendee(const Recipient &recipient) :
		Recipient(recipient)
	{
	}

	bool isOrganizer()
	{
		return ((flags & 0x0000002) != 0);
	}

	void setOrganizer(bool organizer)
	{
		if (organizer) {
			flags |= 0x0000002;
		} else {
			flags &= ~0x0000002;
		}
	}
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

class GalMember {
public:
	QString id;
	QString name;
	QString nick;
	QString email;
	QString title;
	QString organization;
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

protected:
	TALLOC_CTX *m_ctx;

	/**
	 * Debug and error reporting. Each subclass should reimplement with 
	 * logic that emits a prefix identifying the object involved. 
	 */
	virtual QDebug debug() const = 0;
	virtual QDebug error() const = 0;

	/**
	 * A couple of helper functions that sbclasses can use to implement the
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

	bool fetchGAL(QList<GalMember>& list);

	mapi_object_t *d()
	{
		return &m_store;
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
	MapiObject(MapiConnector2 *connection, const char *tallocName, mapi_id_t id);

	virtual ~MapiObject();

	mapi_object_t *d() const;

	mapi_id_t id() const;

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
	 * Fetch a set of properties.
	 * 
	 * @return Whether the pull succeeds, irrespective of whether the tags
	 * were matched.
	 */
	bool propertiesPull(QVector<int> &tags);

	/**
	 * Fetch all properties.
	 */
	virtual bool propertiesPull();

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
	const mapi_id_t m_id;
	struct SPropValue *m_properties;
	uint32_t m_propertyCount;
	mutable mapi_object_t m_object;

private:
	/**
	 * Add a property with the given value, using an immediate assignment.
	 */
	bool propertyWrite(int tag, void *data, bool idempotent = true);
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
	MapiItem(mapi_id_t id, QString &name, QDateTime &modified);

	/**
	 * The id of the full item.
	 */
	mapi_id_t id() const;

	/**
	 * The name of this item.
	 */
	QString name() const;

	/**
	 * The last-modified date time of the full item.
	 */
	QDateTime modified() const;

private:
	const mapi_id_t m_id;
	const QString m_name;
	const QDateTime m_modified;
};

/**
 * Represents a MAPI folder.
 * 
 * @ref MapiItem
 */
class MapiFolder : public MapiObject
{
public:
	MapiFolder(MapiConnector2 *connection, const char *tallocName, mapi_id_t id);

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
	MapiMessage(MapiConnector2 *connection, const char *tallocName, mapi_id_t folderId, mapi_id_t id);

	virtual bool open();

	/**
	 * How many recipients do we have?
	 */
	unsigned recipientCount() const;

	/**
	 * Fetch all recipients.
	 */
	bool recipientsPull();

	/**
	 * Fetch a property by index.
	 */
	Recipient recipientAt(unsigned i) const;

protected:
	const mapi_id_t m_folderId;
	SRowSet m_recipients;

private:
	virtual QDebug debug() const;
	virtual QDebug error() const;
};

/**
 * An Appointment, with attendee recipients.
 */
class MapiAppointment : public MapiMessage
{
public:
	MapiAppointment(MapiConnector2 *connection, const char *tallocName, mapi_id_t folderId, mapi_id_t id);

	/**
	 * Fetch all calendar properties.
	 */
	virtual bool propertiesPull();

	/**
	 * Update a calendar item.
	 */
	virtual bool propertiesPush();

	QString title;
	QString text;
	QString location;
	QString sender;
	QDateTime created;
	QDateTime begin;
	QDateTime end;
	QDateTime modified;
	bool reminderActive;
	QDateTime reminderTime;
	uint32_t reminderMinutes;
	/**
	 * The attendees include not just the recipients (@ref recipientsPull)
	 * but also whoever the meeting was sent-on-behalf-of. That might or 
	 * might not be the sender of the invitation.
	 */
	QList<Attendee> attendees;
	MapiRecurrencyPattern recurrency;

private:
	bool debugRecurrencyPattern(RecurrencePattern *pattern);

	void addUniqueAttendee(Attendee candidate);

	static unsigned isGoodEmailAddress(QString &email);

	virtual QDebug debug() const;
	virtual QDebug error() const;
};

#endif // MAPICONNECTOR2_H
