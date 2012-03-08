/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2012 Shaheed Haque <srhaque@theiet.org>.
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

#ifndef CHANGEPASSWORD_H
#define CHANGEPASSWORD_H

#include <QDialog>

#include "ui_changepassworddialog.h"

class ChangePasswordDialog : public QDialog, public Ui::ChangePasswordDialogBase
{
Q_OBJECT
public:
	explicit ChangePasswordDialog(QWidget* parent = 0, Qt::WindowFlags f = 0);
	virtual ~ChangePasswordDialog() {}

	void setProfileName(QString value);

	QString oldPassword() const;

	QString newPassword() const;

private slots:
	void slotValidateData();
};

#endif // CHANGEPASSWORD_H
