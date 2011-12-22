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

#ifndef EXMAILRESOURCE_H
#define EXMAILRESOURCE_H

#include <mapiresource.h>

class ExMailResource : public MapiResource
{
Q_OBJECT

public:
	ExMailResource(const QString &id);
	virtual ~ExMailResource();

public Q_SLOTS:
	virtual void configure(WId windowId);

protected Q_SLOTS:
	void retrieveCollections();
	void retrieveItems(const Akonadi::Collection &collection);
	bool retrieveItem(const Akonadi::Item &itemOrig, const QSet<QByteArray> &parts);

protected:
	virtual void aboutToQuit();
	virtual void itemAdded(const Akonadi::Item &item, const Akonadi::Collection &collection);
	virtual void itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts);
	virtual void itemRemoved(const Akonadi::Item &item);

private Q_SLOTS:
	/**
	 * Completion handler for itemChanged().
	 */
	void itemChangedContinue(KJob* job);

private:
	bool retrieveAttachments(MapiMessage *message);
};

#endif
