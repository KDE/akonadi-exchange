/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2012-2013 S.R.Haque <srhaque@theiet.org>
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

#include "modifyprofiledialog.h"
#include <QPushButton>
#include <QDebug>

ModifyProfileDialog::ModifyProfileDialog(QWidget* parent, Qt::WindowFlags f): QDialog(parent, f)
{
    setupUi(this);

    connect(lePassword, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
    connect(leUsername, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
    connect(leDomain, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
    connect(leServer, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));

    slotValidateData();
}

void ModifyProfileDialog::slotValidateData()
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

void ModifyProfileDialog::setProfileName(QString value)
{
    leProfile->setText(value);
}

QString ModifyProfileDialog::profileName() const
{
    return leProfile->text();
}

void ModifyProfileDialog::setAttributes(QString &server, QString &domain, QString &username)
{
    leServer->setText(server);
    leDomain->setText(domain);
    leUsername->setText(username);
}

QString ModifyProfileDialog::username() const
{
    return leUsername->text();
}

QString ModifyProfileDialog::password() const
{
    return lePassword->text();
}

QString ModifyProfileDialog::server() const
{
    return leServer->text();
}

QString ModifyProfileDialog::domain() const
{
    return leDomain->text();
}

#include "modifyprofiledialog.moc"
