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

#include "createprofiledialog.h"
#include <QPushButton>
#include <QTimer>
#include <QDebug>

CreateProfileDialog::CreateProfileDialog(QWidget* parent, Qt::WindowFlags f): QDialog(parent, f)
{
	setupUi(this);

	connect(leProfile, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
	connect(leUsername, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
	connect(lePassword, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
	connect(leDomain, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
	connect(leServer, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));

	QTimer::singleShot(0, this, SLOT(slotValidateData()));
}

void CreateProfileDialog::slotValidateData()
{
	bool valid = false;

	if (!leProfile->text().isEmpty() && 
		!leUsername->text().isEmpty() && 
		!lePassword->text().isEmpty() && 
		!leDomain->text().isEmpty() && 
		!leServer->text().isEmpty()) {
		valid = true;
	}

	buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

QString CreateProfileDialog::getProfileName()
{
	return leProfile->text();
}
QString CreateProfileDialog::getUsername()
{
	return leUsername->text();
}
QString CreateProfileDialog::getPassword()
{
	return lePassword->text();
}
QString CreateProfileDialog::getServer()
{
	return leServer->text();
}
QString CreateProfileDialog::getDomain()
{
	return leDomain->text();
}

#include "createprofiledialog.moc"