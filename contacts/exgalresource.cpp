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

#include "exgalresource.h"

#include "settings.h"
#include "settingsadaptor.h"

#include <QtDBus/QDBusConnection>

#include <KLocalizedString>
#include <KABC/Addressee>
#include <KWindowSystem>

#include "mapiconnector2.h"
#include "profiledialog.h"

using namespace Akonadi;

exgalResource::exgalResource( const QString &id )
  : ResourceBase( id ),
  connector( 0 )
{
	new SettingsAdaptor( Settings::self() );
	QDBusConnection::sessionBus().registerObject( QLatin1String( "/Settings" ),
							Settings::self(), QDBusConnection::ExportAdaptors );
}

exgalResource::~exgalResource()
{
	logoff();
}

void exgalResource::retrieveCollections()
{
	kDebug() << "retrieveCollections() called";

	Collection::List collections;

	QStringList folderMimeType;
	folderMimeType << QString::fromAscii("text/calendar");
	folderMimeType << QString::fromAscii("application/x-vnd.kde.contactgroup");

	// create the new collection
	Collection root;
	root.setParentCollection(Collection::root());
	root.setContentMimeTypes(folderMimeType);
	root.setRemoteId(QString::fromAscii("Exchange GAL"));
	root.setName(i18n("Exchange Global Address List"));

	collections.append(root);

	// notify akonadi about the new collection
	collectionsRetrieved(collections);
}

void exgalResource::retrieveItems( const Akonadi::Collection &collection )
{
	Q_UNUSED( collection );

	if (!logon()) {
		return;
	}

	Item::List items;

	QList<GalMember> list;
	emit status(Running, i18n("Fetching GAL from Exchange"));
	if (connector->fetchGAL(list)) {
		int idx = 1;
		foreach (GalMember data, list) {
			Item item(QString::fromAscii("text/directory"));
			item.setParentCollection(collection);
			//item.setRemoteId(data.nick); // find a better remote id (but entryid seams to be binary)
			item.setRemoteId(QString::number(idx++));
			item.setRemoteRevision(QString::number(1));

			// prepare payload
			KABC::Addressee addressee;
			addressee.setName(data.name);
			addressee.setNickName(data.nick);
			addressee.setOrganization(data.organization);

			QStringList emails;
			emails << data.email;
			addressee.setEmails(emails);

			item.setPayload<KABC::Addressee>( addressee );
			items << item;
		}
	}

	itemsRetrieved(items);

	// This seems like a good place to force any subsequent activity
	// to attempt the login.
	logoff();
}

bool exgalResource::retrieveItem( const Akonadi::Item &item, const QSet<QByteArray> &parts )
{
  Q_UNUSED( item );
  Q_UNUSED( parts );

  // TODO: this method is called when Akonadi wants more data for a given item.
  // You can only provide the parts that have been requested but you are allowed
  // to provide all in one go

  return true;
}

void exgalResource::aboutToQuit()
{
  // TODO: any cleanup you need to do while there is still an active
  // event loop. The resource will terminate after this method returns
}

void exgalResource::configure( WId windowId )
{
	ProfileDialog dlgConfig(Settings::self()->profileName());
  	if (windowId)
		KWindowSystem::setMainWindow(&dlgConfig, windowId);

	if (dlgConfig.exec() == KDialog::Accepted) {

		QString profile = dlgConfig.getProfileName();
		Settings::self()->setProfileName( profile );
		Settings::self()->writeConfig();

		synchronize();
		emit configurationDialogAccepted();
	} else {
		emit configurationDialogRejected();
	}
}

void exgalResource::itemAdded( const Akonadi::Item &item, const Akonadi::Collection &collection )
{
  Q_UNUSED( item );
  Q_UNUSED( collection );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has created an item in a collection managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

void exgalResource::itemChanged( const Akonadi::Item &item, const QSet<QByteArray> &parts )
{
  Q_UNUSED( item );
  Q_UNUSED( parts );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has changed an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

void exgalResource::itemRemoved( const Akonadi::Item &item )
{
  Q_UNUSED( item );

  // TODO: this method is called when somebody else, e.g. a client application,
  // has deleted an item managed by your resource.

  // NOTE: There is an equivalent method for collections, but it isn't part
  // of this template code to keep it simple
}

bool exgalResource::logon(void)
{
	if (!connector) {
		// logon to exchange (if needed)
		emit status(Running, i18n("Logging in to Exchange"));
		connector = new MapiConnector2;
	}
	bool ok = connector->login(Settings::self()->profileName());
	if (!ok) {
		emit status(Broken, i18n("Unable to login") );
	}
	return ok;
}

void exgalResource::logoff(void)
{
	delete connector;
	connector = 0;
}

AKONADI_RESOURCE_MAIN( exgalResource )

#include "exgalresource.moc"
