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

#include <QList>
#include <QMap>
#include <QString>
#include <QDateTime>
#include <QList>
#include <QBitArray>

extern "C" {
// libmapi is a C library and must therefore be included that way
// otherwise we'll get linker errors due to C++ name mangling
#include <libmapi/libmapi.h>
}

class CalendarDataShort
{
public:
	QString fid;
	QString id;
	QString title;
	QDateTime modified;
};

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

class CalendarData
{
public:
	QString fid;
	QString id;
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
	uint32_t reminderDelta; // stored in minutes
	QList<Attendee> attendees;
	MapiRecurrencyPattern recurrency;
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

	~TallocContext();

	TALLOC_CTX *d();

private:
	TALLOC_CTX *m_ctx;
};

/**
 * A class for managing access to MAPI profiles. This is the root for all other
 * MAPI interactions.
 */
class MapiProfiles : private TallocContext
{
public:
	MapiProfiles();
	~MapiProfiles();

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

	bool fetchFolderContent(mapi_id_t folderID, QList<CalendarDataShort>& list);
	bool fetchCalendarData(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data);

	/**
	 * Attempt to resolve a listed subset of the given recipients.
	 */
	bool resolveNames(QList<Attendee> &recipients, const QList<unsigned> &needingResolution);

	bool calendarDataUpdate(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data);

	bool fetchGAL(QList<GalMember>& list);

	mapi_object_t *d()
	{
		return &m_store;
	}

private:
	mapi_object_t openFolder(mapi_id_t folderID);

	mapi_session *m_session;
	mapi_object_t m_store;
};

/**
 * A class which wraps a MAPI object such that objects of this type 
 * automatically free the used memory on destruction.
 */
class MapiObject
{
	/**
	 * Add a property with the given value, using an immediate assignment.
	 */
	bool propertyWrite(int tag, void *data, bool idempotent = true);

public:
	MapiObject(TallocContext &ctx, mapi_id_t id);

	virtual ~MapiObject();

	mapi_object_t *d() const;

	mapi_id_t id() const;

	virtual bool open(mapi_object_t *store, mapi_id_t folderId) = 0;

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
	bool propertiesPush();

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
	bool propertiesPull();

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
	TallocContext &m_ctx;
	const mapi_id_t m_id;
	struct SPropValue *m_properties;
	uint32_t m_propertyCount;
	mutable mapi_object_t m_object;
};

/**
 * Represents a MAPI folder.
 */
class MapiFolder : public MapiObject
{
public:
	MapiFolder(TallocContext &ctx, mapi_id_t id);

	virtual ~MapiFolder();

	virtual bool open(mapi_object_t *store, mapi_id_t unused = 0);

	QString id() const;

	QString name;

	/**
	 * Fetch children.
	 * 
	 * @param children The children will be added to this list.
	 * @param filter Only return items whose PR_CONTAINERR_CLASS start this
	 * 		 value.
	 */
	bool childrenPull(QList<MapiFolder> &children, const QString &filter = QString());

protected:
	mapi_object_t m_hierarchyTable;
};

/**
 * A Message, with recipients.
 */
class MapiMessage : public MapiObject
{
public:
	MapiMessage(TallocContext &ctx, mapi_id_t id);

	virtual bool open(mapi_object_t *store, mapi_id_t folderId);

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
	SRowSet m_recipients;
};

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
	 * To convert this into a name, @see MapiObject::tagName().
	 */
	int tag() const;

private:
	SPropValue &m_property;
};

/**
 * An Appointment, with attendee recipients.
 */
class MapiAppointment : public MapiMessage
{
public:
	MapiAppointment(TallocContext &ctx, mapi_id_t id);

	bool open(mapi_object_t *store, mapi_id_t folderId);

	RecurrencePattern *recurrance();

	/**
	 * Fetch a property by index.
	 */
	Attendee recipientAt(unsigned i) const;

	bool getAttendees(QList<Attendee> &attendees, QList<unsigned> &needingResolution);

private:
	bool debugRecurrencyPattern(RecurrencePattern *pattern);
};

#endif // MAPICONNECTOR2_H
