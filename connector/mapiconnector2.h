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

class FolderData
{
public:
	QString id;
	QString name;
};

class CalendarDataShort
{
public:
	QString fid;
	QString id;
	QString title;
	QDateTime modified;
};

class Attendee {
public:
	uint32_t idx;
	QString name;
	QString email;
	bool isOranizer;
	uint32_t status;
	uint32_t type;
};

class MapiRecurrencyPattern {
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
	QList<Attendee> anttendees;
	MapiRecurrencyPattern recurrency;
};

class RecipientData {
public:
	QString name;
	QString email;
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

	~MapiObject();

	mapi_object_t *d();

	mapi_id_t id();

	bool open(mapi_object_t *store, mapi_id_t folderId);

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
	 * Fetch a set of properties.
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
	QString tagAt(unsigned i);

	/**
	 * For display purposes, convert a property into a string, taking
	 * care to hex-ify GUIDs and other byte arrays, and lists of the
	 * same.
	 */
	QString propertyString(unsigned i) const;

	unsigned propertyCount() const;

	/**
	 * Find the name for a tag. If it not a well known one, try a lookup.
	 * Technically, this should only be needed if bit 31 is set, but
	 * still...
	 */
	QString tagName(int tag);

protected:
	TallocContext &m_ctx;
	mapi_id_t m_id;
	struct SPropValue *m_properties;
	uint32_t m_propertyCount;
	mapi_object_t m_object;
};

class MapiAppointment : public MapiObject
{
public:
	MapiAppointment(TallocContext &ctx, mapi_id_t id);

	bool open(mapi_object_t *store, mapi_id_t folderId);

	RecurrencePattern *recurrance();
private:
	bool debugRecurrencyPattern(RecurrencePattern *pattern);
};

class MapiConnector2
{
public:
    MapiConnector2();
    virtual ~MapiConnector2();

	QString getMapiProfileDirectory();

	bool login(QString profilename);

	QStringList listProfiles();
	bool createProfile(QString profile, QString username, QString password, QString domain, QString server);
	bool setDefaultProfile(QString profile);
	QString getDefaultProfile();
	bool removeProfile(QString profile);

	bool fetchFolderList(QList<FolderData>& list, mapi_id_t parentFolderID=0x0, const QString filter=QString());
	bool fetchFolderContent(mapi_id_t folderID, QList<CalendarDataShort>& list);
	bool fetchCalendarData(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data);
	void resolveNames(const QStringList& names, QMap<QString, RecipientData>& outputMap);

	bool calendarDataUpdate(mapi_id_t folderID, mapi_id_t messageID, CalendarData& data);

	bool fetchGAL(QList<GalMember>& list);

	mapi_object_t *d()
	{
		return &m_store;
	}

private:
	mapi_object_t openFolder(mapi_id_t folderID);
	QString mapiValueToQString(mapi_SPropValue *lpProps);
	bool getAttendees(mapi_object_t& obj_message, const QString& toAttendeesStr, CalendarData& data);

	struct mapi_context *m_context;
	mapi_session *m_session;
	mapi_object_t  m_store;
};

#endif // MAPICONNECTOR2_H
