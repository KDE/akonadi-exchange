/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011 Shaheed Haque <srhaque@theiet.org>.
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

#ifndef MAPIRESOURCE_H
#define MAPIRESOURCE_H

#include <akonadi/resourcebase.h>
#include "mapiconnector2.h"

class MapiConnector2;
class MapiFolder;
class MapiMessage;

/**
 * The purpose of this class is to actas a base for individual resources which
 * implement MAPI services. It hides the networking/logon and other details
 * as much as possible.
 */
class MapiResource :
	public Akonadi::ResourceBase,
	public Akonadi::AgentBase::Observer
{
Q_OBJECT

public:
	MapiResource(const QString &id);
	virtual ~MapiResource();

protected:

	/**
	 * Establish the name of the profile to use to logon to the MAPI server.
	 */
	void profileSet(const QString &profileName);

	/**
	 * Recursively find all folders starting at the given root which match
	 * the given filter (i.e. an IPF_xxx value, which matches a 
	 * PidTagContainerClass, or null).
	 * 
	 * @param rootFolder 	Identifies where to start the search.
	 * @param filter	The the needed filter. Set to the empty string 
	 * 			if no filtering is needed.
	 * @param collections	List to which any matches are to be appended.
	 */
	void fetchCollections(MapiDefaultFolder rootFolder, const QString &filter, Akonadi::Collection::List &collections);

	/**
	 * Find all the items in the given collection.
	 * 
	 * @param collection	The collection to fetch.
	 * @param items		Fetched items.
	 * @param deletedItems	Items which have been deleted on the backend.
	 */
	void fetchItems(const Akonadi::Collection &collection, Akonadi::Item::List &items, Akonadi::Item::List &deletedItems);

	/**
	 * Get the message corresponding to the item.
	 */
	template <class Message>
	Message *fetchItem(const Akonadi::Item &item);

protected:
	/*
	virtual void aboutToQuit();
	virtual void itemAdded(const Akonadi::Item &item, const Akonadi::Collection &collection);
	virtual void itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts);
	virtual void itemRemoved(const Akonadi::Item &item);
*/
	/**
	 * Recurse through a hierarchy of Exchange folders which match the
	 * given filter.
	 */
	void fetchCollections(const QString &path, const Akonadi::Collection &parent, const QString &filter, Akonadi::Collection::List &collections);

	/**
	 * Consistent error handling for task-based routines.
	 */
	void error(const QString &message);
	void error(const MapiFolder &folder, const QString &body);
	void error(const Akonadi::Collection &collection, const QString &body);
	void error(const MapiMessage &msg, const QString &body);

	MapiConnector2 *m_connection;
	bool m_connected;

private:
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

	QString m_profileName;
};

#endif
