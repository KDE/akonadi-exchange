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

#include "createprofiledialog.h"

ProfileDialog::ProfileDialog(QString selectedProfile, QWidget* parent)
 : QDialog(parent), selectedProfile(selectedProfile)
{
	setupUi(this);

	connect(btnCreate, SIGNAL(clicked()), this, SLOT(slotCreateNewProfile()));
	connect(btnRemove, SIGNAL(clicked()), this, SLOT(slotRemoveProfile()));

	connect(listWidget, SIGNAL(currentItemChanged(QListWidgetItem*, QListWidgetItem*)), 
			this, SLOT(newProfileSelected(QListWidgetItem*, QListWidgetItem*)));

	QTimer::singleShot(0, this, SLOT(readMapiProfiles()));
	QTimer::singleShot(0, this, SLOT(slotValidate()));
}


void ProfileDialog::readMapiProfiles()
{
	listWidget->clear();

	QStringList profiles = m_profiles.list();

	bool profileExist = false;
	foreach (QString profile, profiles) {
		QListWidgetItem* item = new QListWidgetItem(profile, listWidget);
		if (selectedProfile == profile) {
			listWidget->setCurrentItem(item);
			profileExist = true;
		}
	}

	if (!profileExist) 
		selectedProfile.clear();

	updateSelectedProfile();
}

void ProfileDialog::updateSelectedProfile()
{
	labelSelectedProfile->setText(i18n("Selected Profile:")+QString::fromLatin1(" ")+selectedProfile);
	slotValidate();
}

void ProfileDialog::newProfileSelected(QListWidgetItem* newItem, QListWidgetItem* lastItem)
{
	Q_UNUSED(lastItem)

	if (newItem != NULL) {
		selectedProfile = newItem->text();
	} else {
		selectedProfile.clear();
	}

	updateSelectedProfile();
}

void ProfileDialog::slotValidate()
{
	bool valid = false;

	if (!selectedProfile.isEmpty()) {
		valid = true;
	}

	buttonBox->button(QDialogButtonBox::Ok)->setEnabled(valid);
}

void ProfileDialog::slotCreateNewProfile()
{
	CreateProfileDialog dlg(this);
	if (dlg.exec() == QDialog::Accepted) {
		bool ok = m_profiles.add(dlg.getProfileName(), dlg.getUsername(), dlg.getPassword(), dlg.getDomain(), dlg.getServer());
		if (!ok) 
			KMessageBox::error(this, i18n("An error occurred during the creation of the new profile"));
		readMapiProfiles();
	}
}

void ProfileDialog::slotRemoveProfile()
{
	if (selectedProfile.isEmpty())
		return;

	if (KMessageBox::questionYesNo(this, i18n("Do you really want to delete the selected profile?")) == KMessageBox::Yes) {
		bool ok = m_profiles.remove(selectedProfile);
		if (!ok) 
			KMessageBox::error(this, i18n("An error occurred during the deletion of the selected profile"));
		readMapiProfiles();
	}
}

#include "profiledialog.moc"
