/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011-13 Shaheed Haque <srhaque@theiet.org>.
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

#include <KLocalizedString>
#include <akonadi/resourcebase.h>

#include "mapiobjects.h"

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
    /**
     * @param folderFilter	Folder filter (i.e. an IPF_xxx value, which 
     * 			matches a PidTagContainerClass). Set to the 
     * 			empty string if no filtering is needed.
     */
    MapiResource(const QString &id, const QString &desktopName, const char *folderFilter, const char *messageType, const QString &itemMimeType);
    virtual ~MapiResource();

protected:

    /**
     * Establish the name of the profile to use to logon to the MAPI server.
     */
    virtual const QString profile() = 0;

    /**
     * Recursively find all folders starting at the given root which match
     * the given filter.
     * 
     * @param rootFolder 	Identifies where to start the search.
     * @param collections	List to which any matches are to be appended.
     */
    void fetchCollections(MapiDefaultFolder rootFolder, Akonadi::Collection::List &collections);

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
    virtual void doSetOnline(bool online);

    /**
     * Recurse through a hierarchy of Exchange folders which match the
     * given filter.
     */
    void fetchCollections(const QString &path, const MapiId &parentId, const Akonadi::Collection &parent, Akonadi::Collection::List &collections);

    /**
     * Consistent error handling for task-based routines.
     */
    void error(const QString &message);
    void error(const MapiFolder &folder, const QString &body);
    void error(const Akonadi::Collection &collection, const QString &body);
    void error(const MapiMessage &msg, const QString &body);

    QString m_mapiFolderFilter;
    QString m_mapiMessageType;
    QString m_itemMimeType;
    MapiConnector2 *m_connection;
    bool m_connected;

protected:
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
};

/**
 * Grrr. Stupid C++ and template instantiation requirements - Ada rules!
 */
template <class Message>
Message *MapiResource::fetchItem(const Akonadi::Item &itemOrig)
{
    kDebug() << "fetch item:" << currentCollection().name() << itemOrig.id() <<
            ", " << itemOrig.remoteId();

    if (!logon()) {
        return 0;
    }

    MapiId remoteId(itemOrig.remoteId());
    Message *message = new Message(m_connection, __FUNCTION__, remoteId);
    if (!message->open()) {
        emit status(Broken, i18n("Unable to open item: %1/%2, %3", currentCollection().name(),
                                 itemOrig.id(), mapiError()));
        return 0;
    }

    // find the remoteId of the item and the collection and try to fetch the needed data from the server
    emit status(Running, i18n("Fetching item: %1/%2", currentCollection().name(), itemOrig.id()));
    if (!message->propertiesPull()) {
        emit status(Broken, i18n("Unable to fetch item: %1/%2, %3", currentCollection().name(),
                                 itemOrig.id(), mapiError()));
        delete message;
        return 0;
    }
    kDebug() << "fetched item:" << itemOrig.remoteId();
    return message;
}

#endif
