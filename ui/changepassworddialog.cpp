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

#include "changepassworddialog.h"
#include <QPushButton>
#include <QDebug>

ChangePasswordDialog::ChangePasswordDialog(QWidget* parent, Qt::WindowFlags f): QDialog(parent, f)
{
    setupUi(this);

    connect(leOldPassword, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));
    connect(lePassword, SIGNAL(textEdited(QString)), this, SLOT(slotValidateData()));

    slotValidateData();
}

void ChangePasswordDialog::slotValidateData()
{
    bool valid = false;

    if (!leProfile->text().isEmpty() && 
        !leOldPassword->text().isEmpty() &&
        !lePassword->text().isEmpty()) {
        valid = true;
    }

    buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

QString ChangePasswordDialog::oldPassword() const
{
    return leOldPassword->text();
}

QString ChangePasswordDialog::newPassword() const
{
    return lePassword->text();
}

void ChangePasswordDialog::setProfileName(QString value)
{
    leProfile->setText(value);
}

#include "changepassworddialog.moc"