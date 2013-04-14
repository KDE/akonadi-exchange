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

#include "profiledialog.h"
#include <QDebug>
#include <KMessageBox>

#include "changepassworddialog.h"
#include "createprofiledialog.h"
#include "modifyprofiledialog.h"

ProfileDialog::ProfileDialog(QWidget *parent) : 
    QDialog(parent)
{
    setupUi(this);

    connect(btnCreate, SIGNAL(clicked()), this, SLOT(slotCreateProfile()));
    connect(btnUpdatePassword, SIGNAL(clicked()), this, SLOT(slotUpdatePassword()));
    connect(btnModify, SIGNAL(clicked()), this, SLOT(slotModifyProfile()));
    connect(btnDelete, SIGNAL(clicked()), this, SLOT(slotDeleteProfile()));

    connect(kcfg_ProfileName, SIGNAL(currentItemChanged(QListWidgetItem *, QListWidgetItem *)), 
            this, SLOT(newProfileSelected(QListWidgetItem *, QListWidgetItem *)));

    readMapiProfiles();
}

void ProfileDialog::readMapiProfiles(QString selection)
{
    kcfg_ProfileName->clear();

    foreach (QString entry, m_profiles.list()) {
        QListWidgetItem *item = new QListWidgetItem(entry, kcfg_ProfileName);
        if (selection == entry) {
            kcfg_ProfileName->setCurrentItem(item);
        }
    }
    slotValidate();
}

void ProfileDialog::newProfileSelected(QListWidgetItem* newItem, QListWidgetItem* lastItem)
{
    Q_UNUSED(newItem)
    Q_UNUSED(lastItem)

    slotValidate();
}

void ProfileDialog::slotValidate()
{
    bool valid = false;

    if (!kcfg_ProfileName->currentItem()) {
        goto DONE;
    }
    valid = true;
DONE:
    btnModify->setEnabled(valid);
    btnUpdatePassword->setEnabled(valid);
    btnDelete->setEnabled(valid);
    buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

void ProfileDialog::slotCreateProfile()
{
    CreateProfileDialog dlg(this);

    // Loop until the user gives up, or succeeds!
    while (dlg.exec() == QDialog::Accepted) {
        bool ok = m_profiles.add(dlg.profileName(), dlg.username(), dlg.password(), dlg.domain(), dlg.server());
        if (!ok) {
            KMessageBox::error(this, i18n("An error occurred creating %1: %2", dlg.profileName(), mapiError()));
        } else {
            readMapiProfiles(dlg.profileName());
            break;
        }
    }
}

void ProfileDialog::slotUpdatePassword()
{
    ChangePasswordDialog dlg(this);
    dlg.setProfileName(profileName());

    // Loop until the user gives up, or succeeds!
    while (dlg.exec() == QDialog::Accepted) {
        bool ok = m_profiles.updatePassword(profileName(), dlg.oldPassword(), dlg.newPassword());
        if (!ok) {
            KMessageBox::error(this, i18n("An error occurred changing password for %1: %2", profileName(), mapiError()));
        } else {
            break;
        }
    }
}

void ProfileDialog::slotModifyProfile()
{
    ModifyProfileDialog dlg(this);
    QString username;
    QString domain;
    QString server;

    // Initially, the dialog will accept a password. Use it to fetch the
    // attributes of the profile.
    dlg.setProfileName(profileName());
    bool ok = m_profiles.read(profileName(), username, domain, server);
    if (!ok) {
        KMessageBox::error(this, i18n("Error reading profile %1: %2", profileName(), mapiError()));
        return;
    }

    // Set the attributes of the profile, which can then be edited.
    dlg.setAttributes(server, domain, username);

    // Loop until the user gives up, or succeeds!
    while (dlg.exec() == QDialog::Accepted) {
        bool ok = m_profiles.update(dlg.profileName(), dlg.username(), dlg.password(), dlg.domain(), dlg.server());
        if (!ok) {
            KMessageBox::error(this, i18n("Error modifying profile %1: %2", profileName(), mapiError()));
        } else {
            break;
        }
    }
}

void ProfileDialog::slotDeleteProfile()
{
    if (KMessageBox::questionYesNo(this, i18n("Do you really want to delete %1?", profileName())) == KMessageBox::Yes) {
        bool ok = m_profiles.remove(profileName());
        if (!ok)  {
            KMessageBox::error(this, i18n("An error occurred deleting %1: %2", profileName(), mapiError()));
        }
        readMapiProfiles();
    }
}

QString ProfileDialog::profileName() const
{
    return kcfg_ProfileName->currentItem()->text();
}

void ProfileDialog::setProfileName(QString profile)
{
    readMapiProfiles(profile);
}

#include "profiledialog.moc"
