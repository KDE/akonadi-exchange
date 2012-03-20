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

#ifndef EXGALRESOURCE_H
#define EXGALRESOURCE_H

#include "mapiresource.h"

namespace Akonadi
{
    class Collection;
}
class KJob;
class MapiConnector2;

/**
 * This class gives acces both to the Global Address List (aka the GAL or the
 * Public Address Book, or the PAB) as well as the user's own Contacts.
 */
class ExGalResource : public MapiResource
{
    Q_OBJECT
public:
    ExGalResource(const QString &id);
    virtual ~ExGalResource();

    virtual const QString profile();

public Q_SLOTS:
    virtual void configure(WId windowId);

protected Q_SLOTS:
    void retrieveCollectionAttributes(const Akonadi::Collection &collection);
    void retrieveCollections();
    void retrieveItems(const Akonadi::Collection &collection);
    bool retrieveItem(const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts);

protected:
    virtual void aboutToQuit();

    virtual void itemAdded(const Akonadi::Item &item, const Akonadi::Collection &collection);
    virtual void itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts);
    virtual void itemRemoved(const Akonadi::Item &item);

private:
    /**
     * A copy of the collection used for the GAL.
     */
    class MapiGAL *m_gal;
    Akonadi::Item::List m_galItems;
    qint64 m_msExchangeFetch;
    qint64 m_msAkonadiWrite;
    qint64 m_msAkonadiWriteStatus;
    void updateAkonadiBatchStatus(QString lastAddressee = QString());

private Q_SLOTS:
    void fetchExchangeBatch();
    void createAkonadiItem(KJob *job);
    void createAkonadiItemDone(KJob *job);
    void updateAkonadiBatchStatusDone(KJob *job);
};

#endif
