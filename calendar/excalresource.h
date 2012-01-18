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

#include <mapiresource.h>
#include <KCal/Recurrence>

class MapiRecurrencyPattern;

class ExCalResource : public MapiResource
{
Q_OBJECT

public:
	ExCalResource(const QString &id);
	virtual ~ExCalResource();

	virtual const QString profile();

public Q_SLOTS:
	virtual void configure(WId windowId);

protected Q_SLOTS:
	void retrieveCollections();
	void retrieveItems(const Akonadi::Collection &col);
	bool retrieveItem(const Akonadi::Item &item, const QSet<QByteArray> &parts);

protected:
	virtual void aboutToQuit();
	virtual void itemAdded(const Akonadi::Item &item, const Akonadi::Collection &collection);
	virtual void itemChanged(const Akonadi::Item &item, const QSet<QByteArray> &parts);
	virtual void itemRemoved(const Akonadi::Item &item);

private:
	void createKCalRecurrency(KCal::Recurrence* rec, const MapiRecurrencyPattern& pattern);

private Q_SLOTS:
	/**
	 * Completion handler for itemChanged().
	 */
	void itemChangedContinue(KJob* job);
};

#endif
