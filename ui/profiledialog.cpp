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
#include <QTimer>
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

	connect(profileList, SIGNAL(currentItemChanged(QListWidgetItem*, QListWidgetItem*)), 
			this, SLOT(newProfileSelected(QListWidgetItem*, QListWidgetItem*)));

	QTimer::singleShot(0, this, SLOT(readMapiProfiles()));
}


void ProfileDialog::readMapiProfiles()
{
	profileList->clear();

	foreach (QString entry, m_profiles.list()) {
		QListWidgetItem* item = new QListWidgetItem(entry, profileList);
		if (kcfg_ProfileName->text() == entry) {
			profileList->setCurrentItem(item);
		}
	}
	updateSelectedProfile();
}

void ProfileDialog::updateSelectedProfile()
{
	if (profileList->currentItem()) {
		kcfg_ProfileName->setText(profileList->currentItem()->text());
	} else {
		kcfg_ProfileName->clear();
	}
	slotValidate();
}

void ProfileDialog::newProfileSelected(QListWidgetItem* newItem, QListWidgetItem* lastItem)
{
	Q_UNUSED(newItem)
	Q_UNUSED(lastItem)

	updateSelectedProfile();
}

void ProfileDialog::slotValidate()
{
	bool valid = false;

	if (kcfg_ProfileName->text().isEmpty()) {
		goto DONE;
	}
	valid = true;
DONE:
	buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

void ProfileDialog::slotCreateProfile()
{
	CreateProfileDialog dlg(this);
	if (dlg.exec() == QDialog::Accepted) {
		bool ok = m_profiles.add(dlg.profileName(), dlg.username(), dlg.password(), dlg.domain(), dlg.server());
		if (!ok) {
			KMessageBox::error(this, i18n("An error occurred during the creation of the new profile"));
		} else {
			kcfg_ProfileName->setText(dlg.profileName());
		}
		readMapiProfiles();
	}
}

void ProfileDialog::slotUpdateProfile()
{
	ChangePasswordDialog dlg(this);
	dlg.setProfileName(kcfg_ProfileName->text());
	if (dlg.exec() == QDialog::Accepted) {
		bool ok = m_profiles.updatePassword(kcfg_ProfileName->text(), dlg.oldPassword(), dlg.newPassword());
		if (!ok) {
			KMessageBox::error(this, i18n("An error occurred during the changing of the password: %1", mapiError()));
		}
		readMapiProfiles();
	}
}

void ProfileDialog::slotDeleteProfile()
{
	if (kcfg_ProfileName->text().isEmpty())
		return;

	if (KMessageBox::questionYesNo(this, i18n("Do you really want to delete the selected profile?")) == KMessageBox::Yes) {
		bool ok = m_profiles.remove(selectedProfile);
		if (!ok)  {
			KMessageBox::error(this, i18n("An error occurred during the deletion of the selected profile"));
		}
		readMapiProfiles();
	}
}

QString ProfileDialog::profileName() const
{
	return kcfg_ProfileName->text();
}

void ProfileDialog::setProfileName(QString profile)
{
	kcfg_ProfileName->setText(profile);
	readMapiProfiles();
}


#include "profiledialog.moc"
