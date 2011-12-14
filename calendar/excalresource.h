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

#ifndef EXCALRESOURCE_H
#define EXCALRESOURCE_H

#include <akonadi/resourcebase.h>
#include <KCal/Recurrence>

class MapiConnector2;
class MapiFolder;
class MapiMessage;
class MapiRecurrencyPattern;

class ExCalResource : public Akonadi::ResourceBase,
                           public Akonadi::AgentBase::Observer
{
Q_OBJECT

public:
	ExCalResource( const QString &id );
	~ExCalResource();

public Q_SLOTS:
	virtual void configure( WId windowId );

protected Q_SLOTS:
	void retrieveCollections();
	void retrieveItems( const Akonadi::Collection &col );
	bool retrieveItem( const Akonadi::Item &item, const QSet<QByteArray> &parts );

protected:
	virtual void aboutToQuit();
	virtual void itemAdded( const Akonadi::Item &item, const Akonadi::Collection &collection );
	virtual void itemChanged( const Akonadi::Item &item, const QSet<QByteArray> &parts );
	virtual void itemRemoved( const Akonadi::Item &item );

private:
	void createKCalRecurrency(KCal::Recurrence* rec, const MapiRecurrencyPattern& pattern);

	/**
	 * Logon to Exchange. A successful login is cached and subsequent calls
	 * short-circuited.
	 *
	 * @return True if the login attempt succeeded.
	 */
	bool logon(void);

	/**
	 * Logout from Exchange.
	 */
	void logoff(void);

	MapiConnector2 *m_connection;
	bool m_connected;

	/**
	 * Consistent error handling for task-based routines.
	 */
	void error(const QString &message);
	void error(const MapiFolder &folder, const QString &body);
	void error(const Akonadi::Collection &collection, const QString &body);
	void error(const MapiMessage &msg, const QString &body);

private Q_SLOTS:
	/**
	 * Completion handler for itemChanged().
	 */
	void itemChangedContinue(KJob* job);
};

#endif
