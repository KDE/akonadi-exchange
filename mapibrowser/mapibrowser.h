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

#ifndef mapibrowser_H
#define mapibrowser_H

#include <QtGui/QMainWindow>

#include "mainwindow.h"


class mapibrowser : public QMainWindow
{
Q_OBJECT
public:
    mapibrowser();
    virtual ~mapibrowser() {}

public slots:
	void onRefreshTree();
	void itemDoubleClicked(QTreeWidgetItem*,int);
	void itemDoubleClicked(QTableWidgetItem*);
	void checkForDefaultProfile();

private:
	MainWindow* main;
	QString selectedProfile;
};

#endif // mapibrowser_H
