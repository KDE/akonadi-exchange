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

#ifndef __EXCALCONFIG_H__
#define __EXCALCONFIG_H__

#include <QDialog>
#include "ui_profiledialog.h"
#include "mapiconnector2.h"

class ProfileDialog : public QDialog, public Ui::ProfileDialogBase
{
Q_OBJECT
public:
	explicit ProfileDialog(QString selectedProfile, QWidget *parent = 0);
	virtual ~ProfileDialog() {}

	QString getProfileName() { return selectedProfile; }

private slots:
	void readMapiProfiles();
	void newProfileSelected(QListWidgetItem* newItem, QListWidgetItem* lastItem);
	void slotValidate();

	void slotCreateNewProfile();
	void slotRemoveProfile();

private:
	void updateSelectedProfile();

	MapiProfiles m_profiles;
	QString selectedProfile;
};

#endif
