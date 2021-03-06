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

#include "mapibrowser.h"

#include <QDebug>
#include <QtGui/QLabel>
#include <QtGui/QMenu>
#include <QtGui/QMenuBar>
#include <QtGui/QAction>
#include <QtGui/QMessageBox>
#include <QtGui/QListWidgetItem>
#include <QInputDialog>
#include <QTimer>

#include "mapiobjects.h"
#include "profiledialog.h"

mapibrowser::mapibrowser()
{
    main = new MainWindow(this);
    setCentralWidget( main );

    QMenu* fileMenu = menuBar()->addMenu( QString::fromLocal8Bit("File") );

    QAction* a = new QAction(this);
    a->setText( QString::fromLocal8Bit("Get Folder Tree") );
    connect(a, SIGNAL(triggered()), SLOT(onRefreshTree()) );
    fileMenu->addAction( a );

    a = new QAction(this);
    a->setText( QString::fromLocal8Bit("Manage Profiles") );
    connect(a, SIGNAL(triggered()), SLOT(onManageProfiles()) );
    fileMenu->addAction( a );

    a = new QAction(this);
    a->setText( QString::fromLocal8Bit("Quit") );
    connect(a, SIGNAL(triggered()), SLOT(close()) );
    fileMenu->addAction( a );


    // connect signals from the generated UI
    connect(main->treeWidget, SIGNAL(itemDoubleClicked(QTreeWidgetItem*,int)), 
            this, SLOT(itemDoubleClicked(QTreeWidgetItem*,int)));
    connect(main->tableWidget, SIGNAL(itemDoubleClicked(QTableWidgetItem*)), 
            this, SLOT(itemDoubleClicked(QTableWidgetItem*)));

    QTimer::singleShot(0, this, SLOT(checkForDefaultProfile()));
}


void mapibrowser::checkForDefaultProfile()
{
    MapiProfiles profiles;
    QString profile = profiles.defaultGet();
    if (profile.isEmpty()) {
        // no default profile -> query the user
        bool ok;
        profile = QInputDialog::getItem(this, QString::fromAscii("Select Profile"), QString::fromAscii("Profile:"), profiles.list(), 0, false, &ok);
        if (!ok || profile.isEmpty()) {
            return;
        }
        this->selectedProfile = profile;
    }
}

void mapibrowser::onRefreshTree()
{
    MapiConnector2 con;
    bool ok = con.login(selectedProfile);
    if (!ok) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("Login failed!"));
        return;
    }

    MapiId rootId(&con, TopInformationStore);
    if (!rootId.isValid())
    {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("cannot find Exchange folder root"));
        return;
    }
    MapiFolder rootFolder(&con, "mapibrowser::onRefreshTree", rootId);
    if (!rootFolder.open()) {
        return;
    }

    QList<MapiFolder *> list;
    if (!rootFolder.childrenPull(list)) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("childrenPull failed!"));
        return;
    }

    QTreeWidgetItem *root = new QTreeWidgetItem(main->treeWidget);
    root->setText(0, QString::fromLocal8Bit("Mailbox"));

    foreach (const MapiFolder *data, list) {
        QTreeWidgetItem *item = new QTreeWidgetItem(root);
        item->setText(0, data->name);
        item->setText(1, data->id().toString());
        delete data;
    }

    root->setExpanded(true);
}

void mapibrowser::onManageProfiles()
{
    ProfileDialog dlgConfig;
    dlgConfig.setProfileName(selectedProfile);

    if (dlgConfig.exec() == QDialog::Accepted) {
        selectedProfile = dlgConfig.profileName();
    }
}

void mapibrowser::itemDoubleClicked(QTreeWidgetItem* clickedItem, int )
{
    QString remoteId = clickedItem->text(1);

    MapiConnector2 con;
    bool ok = con.login(selectedProfile);
    if (!ok) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("Login failed!"));
        return;
    }

    MapiId id(remoteId);
    MapiFolder parentFolder(&con, "mapibrowser::itemDoubleClicked", id);
    if (!parentFolder.open()) {
        return;
    }

    QList<MapiItem *> list;
    if (!parentFolder.childrenPull(list)) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("childrenPull failed!\n")+remoteId);
        return;
    }

    main->tableWidget->clearContents();
    main->tableWidget->setSelectionBehavior(QAbstractItemView::SelectRows);

    main->tableWidget->setRowCount(list.size());
    int row=0;
    foreach (const MapiItem *data, list) {
        QTableWidgetItem *newItem;
        newItem = new QTableWidgetItem(data->id().toString());
        main->tableWidget->setItem(row, 0, newItem);
        newItem = new QTableWidgetItem(data->name());
        main->tableWidget->setItem(row, 1, newItem);
        newItem = new QTableWidgetItem(data->modified().toString());
        main->tableWidget->setItem(row, 2, newItem);
        delete data;
        row++;
    }

    main->tableWidget->setEditTriggers(QTableWidget::NoEditTriggers); 
}

void mapibrowser::itemDoubleClicked(QTableWidgetItem* clickedItem)
{
    QTableWidgetItem* idItem = main->tableWidget->item(clickedItem->row(), 0);
    QString messageId = idItem->text();

    QString folderId = main->treeWidget->selectedItems().at(0)->text(1);

    MapiConnector2 con;
    bool ok = con.login(selectedProfile);
    if (!ok) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("Login failed!"));
        return;
    }

#if 0
// JUST FOR DEBUG --BEGIN--
    MapiAppointment appoint(&con, "itemDoubleClicked-DEBUG",folderId.toULongLong(), messageId.toULongLong());
    if (!appoint.open()) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("open failed!\n")+folderId+QString::fromLocal8Bit(":")+messageId);
        return;
    }
    appoint.propertiesPull();
// JUST FOR DEBUG --END--
#endif
    // For now, hardcode the provider to EMSDB.
    QString remoteId = QString::fromAscii("1/%1/%2").arg(folderId).arg(messageId);
    MapiId id(remoteId);
    MapiMessage message(&con, "mapibrowser::itemDoubleClicked", id);
    if (!message.open()) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("open failed!\n")+folderId+QString::fromLocal8Bit(":")+messageId);
        return;
    }
    if (!message.propertiesPull()) {
        QMessageBox::warning(this, QString::fromLocal8Bit("Error"), QString::fromLocal8Bit("pull failed!\n")+folderId+QString::fromLocal8Bit(":")+messageId);
        return;
    }

    main->tableWidgetDetail->clearContents();
    main->tableWidgetDetail->setSelectionBehavior(QAbstractItemView::SelectRows);
    main->tableWidgetDetail->setRowCount(message.propertyCount());
    for (unsigned row = 0; row < message.propertyCount(); row++) {
        QTableWidgetItem *newItem;

        newItem = new QTableWidgetItem(message.tagAt(row));
        main->tableWidgetDetail->setItem(row, 0, newItem);
        newItem = new QTableWidgetItem(message.propertyString(row));
        main->tableWidgetDetail->setItem(row, 1, newItem);
    }
}

#include "mapibrowser.moc"
