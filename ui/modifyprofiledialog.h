/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2012 S.R.Haque <srhaque@theiet.org>
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

#ifndef MODIFYPROFILE_H
#define MODIFYPROFILE_H

#include <QDialog>

#include "ui_modifyprofiledialog.h"

class ModifyProfileDialog : public QDialog, public Ui::ModifyProfileDialogBase
{
    Q_OBJECT
public:
    explicit ModifyProfileDialog(QWidget* parent = 0, Qt::WindowFlags f = 0);
    virtual ~ModifyProfileDialog() {}

    void setProfileName(QString value);
    QString profileName() const;

    QString password() const;

    void setAttributes(QString &server, QString &domain, QString &username);
    QString server() const;
    QString domain() const;
    QString username() const;

private slots:
    void slotValidateData();
};

#endif // MODIFYPROFILE_H
