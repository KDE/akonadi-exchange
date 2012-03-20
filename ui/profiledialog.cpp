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

ProfileDialog::ProfileDialog(QWidget *parent) : 
    QDialog(parent)
{
    setupUi(this);

    connect(btnCreate, SIGNAL(clicked()), this, SLOT(slotCreateProfile()));
    connect(btnUpdate, SIGNAL(clicked()), this, SLOT(slotUpdateProfile()));
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
    btnUpdate->setEnabled(valid);
    btnDelete->setEnabled(valid);
    buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

void ProfileDialog::slotCreateProfile()
{
    CreateProfileDialog dlg(this);
    if (dlg.exec() == QDialog::Accepted) {
        bool ok = m_profiles.add(dlg.profileName(), dlg.username(), dlg.password(), dlg.domain(), dlg.server());
        if (!ok) {
            KMessageBox::error(this, i18n("An error occurred creating %1: %2", dlg.profileName(), mapiError()));
        }
        readMapiProfiles(dlg.profileName());
    }
}

void ProfileDialog::slotUpdateProfile()
{
    ChangePasswordDialog dlg(this);
    dlg.setProfileName(profileName());
    if (dlg.exec() == QDialog::Accepted) {
        bool ok = m_profiles.updatePassword(profileName(), dlg.oldPassword(), dlg.newPassword());
        if (!ok) {
            KMessageBox::error(this, i18n("An error occurred changing password for %1: %2", mapiError()));
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