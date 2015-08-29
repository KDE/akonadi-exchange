/*
 * This file is part of the Akonadi Exchange Resource.
 * Copyright 2011-13 Robert Gruber <rgruber@users.sourceforge.net>, Shaheed Haque
 * <srhaque@theiet.org>.
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

#include "mapiconnector2.h"

#include <QAbstractSocket>
#include <QStringList>
#include <QDir>
#include <QMessageBox>
#include <QRegExp>
#include <QVariant>
#include <QSocketNotifier>
#include <QTextCodec>
#include <KDebug>
#include <KLocale>
#include <kpimutils/email.h>

#ifndef ENABLE_MAPI_DEBUG
#define ENABLE_MAPI_DEBUG 1
#endif

#ifndef ENABLE_NOTIFICATIONS
#define ENABLE_NOTIFICATIONS 0
#endif

#ifndef ENABLE_PUBLIC_FOLDERS
#define ENABLE_PUBLIC_FOLDERS 0
#endif

#define STR(def) \
case def: return QString::fromLatin1(#def)

/**
 * Map all MAPI errors to strings.
 */
QString mapiError()
{
    int code = GetLastError();
    switch (code)
    {
    STR(MAPI_E_SUCCESS);
    STR(MAPI_E_CALL_FAILED);
    STR(MAPI_E_NO_SUPPORT);
    STR(MAPI_E_BAD_CHARWIDTH);
    STR(MAPI_E_STRING_TOO_LONG);
    STR(MAPI_E_UNKNOWN_FLAGS);
    STR(MAPI_E_INVALID_ENTRYID);
    STR(MAPI_E_INVALID_OBJECT);
    STR(MAPI_E_OBJECT_CHANGED);
    STR(MAPI_E_OBJECT_DELETED);
    STR(MAPI_E_BUSY);
    STR(MAPI_E_NOT_ENOUGH_DISK);
    STR(MAPI_E_NOT_ENOUGH_RESOURCES);
    STR(MAPI_E_NOT_FOUND);
    STR(MAPI_E_VERSION);
    STR(MAPI_E_LOGON_FAILED);
    STR(MAPI_E_SESSION_LIMIT);
    STR(MAPI_E_USER_CANCEL);
    STR(MAPI_E_UNABLE_TO_ABORT);
    STR(ecRpcFailed);
    STR(MAPI_E_DISK_ERROR);
    STR(MAPI_E_TOO_COMPLEX);
    STR(MAPI_E_BAD_COLUMN);
    STR(MAPI_E_EXTENDED_ERROR);
    STR(MAPI_E_COMPUTED);
    STR(MAPI_E_CORRUPT_DATA);
    STR(MAPI_E_UNCONFIGURED);
    STR(MAPI_E_FAILONEPROVIDER);
    STR(MAPI_E_UNKNOWN_CPID);
    STR(MAPI_E_UNKNOWN_LCID);
    STR(MAPI_E_PASSWORD_CHANGE_REQUIRED);
    STR(MAPI_E_PASSWORD_EXPIRED);
    STR(MAPI_E_INVALID_WORKSTATION_ACCOUNT);
    STR(MAPI_E_INVALID_ACCESS_TIME);
    STR(MAPI_E_ACCOUNT_DISABLED);
    STR(MAPI_E_END_OF_SESSION);
    STR(MAPI_E_UNKNOWN_ENTRYID);
    STR(MAPI_E_MISSING_REQUIRED_COLUMN);
    STR(MAPI_E_BAD_VALUE);
    STR(MAPI_E_INVALID_TYPE);
    STR(MAPI_E_TYPE_NO_SUPPORT);
    STR(MAPI_E_UNEXPECTED_TYPE);
    STR(MAPI_E_TOO_BIG);
    STR(MAPI_E_DECLINE_COPY);
    STR(MAPI_E_UNEXPECTED_ID);
    STR(MAPI_E_UNABLE_TO_COMPLETE);
    STR(MAPI_E_TIMEOUT);
    STR(MAPI_E_TABLE_EMPTY);
    STR(MAPI_E_TABLE_TOO_BIG);
    STR(MAPI_E_INVALID_BOOKMARK);
    STR(MAPI_E_WAIT);
    STR(MAPI_E_CANCEL);
    STR(MAPI_E_NOT_ME);
    STR(MAPI_E_CORRUPT_STORE);
    STR(MAPI_E_NOT_IN_QUEUE);
    STR(MAPI_E_NO_SUPPRESS);
    STR(MAPI_E_COLLISION);
    STR(MAPI_E_NOT_INITIALIZED);
    STR(MAPI_E_NON_STANDARD);
    STR(MAPI_E_NO_RECIPIENTS);
    STR(MAPI_E_SUBMITTED);
    STR(MAPI_E_HAS_FOLDERS);
    STR(MAPI_E_HAS_MESAGES);
    STR(MAPI_E_FOLDER_CYCLE);
    STR(MAPI_E_AMBIGUOUS_RECIP);
    STR(MAPI_E_NO_ACCESS);
    STR(MAPI_E_INVALID_PARAMETER);
    STR(MAPI_E_RESERVED);
    default:
        return QString::fromAscii("MAPI_E_0x%1").arg((unsigned)code, 0, 16);
    }
}

static int profileSelectCallback(PropertyRowSet_r *rowset, const void* /*private_var*/)
{
    qCritical() << "Found more than 1 matching users -> cancel";

    //  TODO Some sort of handling would be needed here
    return rowset->cRows;
}

/**
 * An object which exists to enable MAPI debugging. This is intended to be
 * instantiated on the stack, as its destruction automatically resets the
 * debug settings.
 */
class MapiDebug
{
public:
    MapiDebug(MapiProfiles *context, bool enable) :
        m_context(context->m_context),
        m_enabled(enable)
    {
        if (m_enabled) {
            if (MAPI_E_SUCCESS != SetMAPIDebugLevel(m_context, 9)) {
                kError() << "cannot set debug level" << mapiError();
            }
            if (MAPI_E_SUCCESS != SetMAPIDumpData(m_context, true)) {
                kError() << "cannot set dump data" << mapiError();
            }
        }
    }

    ~MapiDebug()
    {
        if (m_enabled) {
            if (MAPI_E_SUCCESS != SetMAPIDebugLevel(m_context, 0)) {
                kError() << "cannot reset debug level" << mapiError();
            }
            if (MAPI_E_SUCCESS != SetMAPIDumpData(m_context, false)) {
                kError() << "cannot reset dump data" << mapiError();
            }
        }
    }
private:
    mapi_context *m_context;
    bool m_enabled;
};

MapiConnector2::MapiConnector2() :
    MapiProfiles(),
    m_session(0),
    m_notifier(0)
{
    m_store = allocate<mapi_object_t>();
    m_nspiStore = allocate<mapi_object_t>();
    mapi_object_init(m_store);
    mapi_object_init(m_nspiStore);
}

MapiConnector2::~MapiConnector2()
{
    delete m_notifier;
    // TODO The calls to tidy up m_nspiStore seem to break things.
    if (m_session) {
        //Logoff(m_nspiStore);
        Logoff(m_store);
    }
    //mapi_object_release(m_nspiStore);
    //mapi_object_release(m_store);
}

QDebug MapiConnector2::debug() const
{
    static QString prefix = QString::fromAscii("MapiConnector2:");
    return TallocContext::debug(prefix);
}

bool MapiConnector2::defaultFolder(MapiDefaultFolder folderType, MapiId *id)
{
    // In case we fail...
    id->m_provider = MapiId::INVALID;

#if (!ENABLE_PUBLIC_FOLDERS)
    error() << "public folders disabled";
    return false;
#endif
    // NSPI-based assets.
    if ((PublicRoot <= folderType) && (folderType <= PublicNNTPArticle)) {
        if (MAPI_E_SUCCESS != GetDefaultPublicFolder(m_nspiStore, &id->second, folderType)) {
            error() << "cannot get default public folder: %1" << folderType << mapiError();
            return false;
        }
        id->first = 0;
        id->m_provider = MapiId::NSPI;
        return true;
    }

    // EMSDB-based assets.
    if (MAPI_E_SUCCESS != GetDefaultFolder(m_store, &id->second, folderType)) {
        error() << "cannot get default folder: %1" << folderType << mapiError();
        return false;
    }
    id->first = 0;
    id->m_provider = MapiId::EMSDB;
    return true;
}

QDebug MapiConnector2::error() const
{
    static QString prefix = QString::fromAscii("MapiConnector2:");
    return TallocContext::error(prefix);
}

bool MapiConnector2::GALCount(unsigned *totalCount)
{
    if (MAPI_E_SUCCESS != GetGALTableCount(m_session, totalCount)) {
        error() << "cannot get GAL count" << mapiError();
        return false;
    }
    return true;
}

bool MapiConnector2::GALRead(unsigned requestedCount, SPropTagArray *tags, SRowSet **results, unsigned *percentagePosition)
{
    if (MAPI_E_SUCCESS != GetGALTable(m_session, tags, (PropertyRowSet_r **)results, requestedCount, TABLE_CUR)) {
        error() << "cannot read GAL entries" << mapiError();
        return false;
    }

    // Return where we got to.
    if (percentagePosition) {
        struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;

        *percentagePosition = nspi->pStat->NumPos * 100 / nspi->pStat->TotalRecs;
    }
    return true;
}

bool MapiConnector2::GALRewind()
{
    struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;

    nspi->pStat->CurrentRec = (NSPI_MID)MID_BEGINNING_OF_TABLE;
    nspi->pStat->Delta = 0;
    nspi->pStat->NumPos = 0;
    nspi->pStat->TotalRecs = 0xffffffff;
    return true;
}

bool MapiConnector2::GALSeek(const QString &displayName, unsigned *percentagePosition, SPropTagArray *tags, SRowSet **results)
{
    struct nspi_context *nspi = (struct nspi_context *)m_session->nspi->ctx;
    SPropValue key;
    SRowSet *dummy;

    key.ulPropTag = (MAPITAGS)PR_DISPLAY_NAME_UNICODE;
    key.dwAlignPad = 0;
    key.value.lpszW = string(displayName);
    if (MAPI_E_SUCCESS != nspi_SeekEntries(nspi, ctx(), SortTypeDisplayName, (PropertyValue_r *)&key, tags, NULL, (PropertyRowSet_r **)(results ? results : &dummy))) {
        error() << "cannot seek to GAL entry" << displayName << mapiError();
        return false;
    }

    // Return where we got to.
    if (percentagePosition) {
        *percentagePosition = nspi->pStat->NumPos * 100 / nspi->pStat->TotalRecs;
    }
    return true;
}

bool MapiConnector2::login(QString profile)
{
    if (!init()) {
        return false;
    }
    if (m_session) {
        // already logged in...
        return true;
    }

    if (profile.isEmpty()) {
        // use the default profile if none was specified by the caller
        profile = defaultGet();
        if (profile.isEmpty()) {
            // there seams to be no default profile
            error() << "no default profile";
            return false;
        }
    }

    // Log on
    if (MAPI_E_SUCCESS != MapiLogonEx(m_context, &m_session, profile.toUtf8(), NULL)) {
        error() << "cannot logon using profile" << profile << mapiError();
        return false;
    }
    if (MAPI_E_SUCCESS != OpenMsgStore(m_session, m_store)) {
        error() << "cannot open message store" << mapiError();
        return false;
    }
#if (ENABLE_PUBLIC_FOLDERS)
    if (MAPI_E_SUCCESS != OpenPublicFolder(m_session, m_nspiStore)) {
        error() << "cannot open public folder" << mapiError();
        return false;
    }
#endif
    MapiDebug debug(this, (0 != ENABLE_MAPI_DEBUG));

    // Get rid of any existing notifier and create a new one.
    // TODO Wait for a version of libmapi that has asingle parameter here.
#if (ENABLE_NOTIFICATIONS)
#if 0
    if (MAPI_E_SUCCESS != RegisterNotification(m_session)) {
#else
    if (MAPI_E_SUCCESS != RegisterNotification(m_session, 0)) {
#endif
        error() << "cannot register for notifications" << mapiError();
        return false;
    }
    delete m_notifier;
    m_notifier = new QSocketNotifier(m_session->notify_ctx->fd, QSocketNotifier::Read);
    if (!m_notifier) {
        error() << "cannot create notifier";
        return false;
    }
    connect(m_notifier, SIGNAL(activated(int)), this, SLOT(notified(int)));
#endif
    return true;
}

void MapiConnector2::notified(int fd)
{
    QAbstractSocket socket(QAbstractSocket::UdpSocket, this);
    socket.setSocketDescriptor(fd);
    struct mapi_response    *mapi_response;
    NTSTATUS                status;
    QByteArray data;
    while (true) {
        data = socket.readAll();
        error() << "read from socket" << data.size();
        if (!data.size()) {
            break;
        }
        // Dummy transaction to keep the  pipe up.
        status = emsmdb_transaction_null((struct emsmdb_context *)m_session->emsmdb->ctx,
                                                                         &mapi_response);
        if (!NT_STATUS_IS_OK(status)) {
            error() << "bad nt status" << status;
            break;
        }
        //retval = ProcessNotification(notify_ctx, mapi_response);
    }
}

bool MapiConnector2::resolveNames(const char *names[], SPropTagArray *tags,
                  SRowSet **results, PropertyTagArray_r **statuses)
{
    if (MAPI_E_SUCCESS != ResolveNames(m_session, names, tags, (PropertyRowSet_r **)results, statuses, MAPI_UNICODE)) {
        error() << "cannot resolve names" << mapiError();
        return false;
    }
    return true;
}

/**
 * We store all objects in Akonadi using the densest string representation to hand:
 *
 *      (0|1|2)/base-36-parentId/base-36-id
 *
 * where the leading 0|1|2 signifies whether the provider is INVALID, EMSDB or NSPI.
 */
const QChar MapiId::fidIdSeparator = QChar::fromAscii('/');

#define ID_BASE 36

MapiId::MapiId(class MapiConnector2 *connection, MapiDefaultFolder folderType)
{
    connection->defaultFolder(folderType, this);
    m_provider = EMSDB;
}

MapiId::MapiId(const MapiId &parent, const mapi_id_t &child)
{
    m_provider = parent.m_provider;
    first = parent.second;
    second = child;
}

MapiId::MapiId(const QString &id)
{
    m_provider = (Provider)id.at(0).digitValue();
    int separator = id.indexOf(fidIdSeparator, 2);
    first = id.mid(2, separator - 2).toULongLong(0, ID_BASE);
    second = id.mid(separator + 1).toULongLong(0, ID_BASE);
}

QString MapiId::toString() const
{
    QChar provider = QChar::fromAscii('0' + m_provider);
    QString parentId = QString::number(first, ID_BASE);
    QString id = QString::number(second, ID_BASE);
    return QString::fromAscii("%1/%2/%3").arg(provider).arg(parentId).arg(id);
}

bool MapiId::isValid() const
{
    return (m_provider == EMSDB) || (m_provider == NSPI);
}

MapiProfiles::MapiProfiles() :
    TallocContext("MapiProfiles::MapiProfiles"),
    m_context(0),
    m_initialised(false)
{
}

MapiProfiles::~MapiProfiles()
{
    if (m_context) {
        MAPIUninitialize(m_context);
    }
}

bool MapiProfiles::add(const QString &profile, const QString &username, const QString &password, const QString &domain, const QString &server)
{
    if (!init()) {
        return false;
    }

    // It seams that the QByteArray returned by QString::toUtf8() are not backed up.
    // calling toUtf8() several times will destroy the data. Therefore I copy them
    // all to temporary objects in order to preserve them throughout the whole method
    QByteArray baProfile(profile.toUtf8());
    QByteArray baUser(username.toUtf8());
    QByteArray baPass(password.toUtf8());
    QByteArray baDomain(domain.toUtf8());
    QByteArray baServer(server.toUtf8());

    const char *profile8 = baProfile.constData();
    qDebug() << "New profile is:"<<profile8;

    if (MAPI_E_SUCCESS != CreateProfile(m_context, profile8, baUser.constData(), baPass.constData(), 0)) {
        error() << "cannot create profile:" << mapiError();
        return false;
    }

    // TODO get workstation as parameter (was is it needed for anyway?)
    char hostname[256] = {};
    gethostname(&hostname[0], sizeof(hostname) - 1);
    hostname[sizeof(hostname) - 1] = 0;
    QString workstation = QString::fromLatin1(hostname);

    if (!attributeAdd(profile8, "binding", server)) {
        return false;
    }
    if (!attributeAdd(profile8, "workstation", workstation)) {
        return false;
    }
    if (!attributeAdd(profile8, "domain", domain)) {
        return false;
    }
// What is seal for? Seams to have something to do with Exchange 2010
// 	mapi_profile_add_string_attr(m_context, profile.toUtf8().constData(), "seal", (seal == true) ? "true" : "false");

// TODO Get langage from parameter if needed
// 	const char* locale = (const char *) (language) ? mapi_get_locale_from_language(language) : mapi_get_system_locale();
    const char *locale = mapi_get_system_locale();
    if (!locale) {
        error() << "cannot find system locale:" << mapiError();
        return false;
    }

    uint32_t cpid = mapi_get_cpid_from_locale(locale);
    uint32_t lcid = mapi_get_lcid_from_locale(locale);
    if (!cpid || !lcid) {
        error() << "invalid Locale supplied or unknown system locale" << locale << ", deleting profile..." << mapiError();
        if (!remove(profile)) {
            return false;
        }
        return false;
    }

    if (!attributeAdd(profile8, "codepage", QString::number(cpid))) {
        return false;
    }
    if (!attributeAdd(profile8, "language", QString::number(lcid))) {
        return false;
    }
    if (!attributeAdd(profile8, "method", QString::number(lcid))) {
        return false;
    }

    struct mapi_session *session = NULL;
    if (MAPI_E_SUCCESS != MapiLogonProvider(m_context, &session, profile8, password.toUtf8(), PROVIDER_ID_NSPI)) {
        error() << "cannot get logon provider, deleting profile..." << mapiError();
        if (!remove(profile)) {
            return false;
        }
        return false;
    }

    int retval = ProcessNetworkProfile(session, username.toUtf8().constData(), profileSelectCallback, NULL);
    if (retval != MAPI_E_SUCCESS && retval != 0x1) {
        error() << "cannot process network profile, deleting profile..." << mapiError();
        if (!remove(profile)) {
            return false;
        }
        return false;
    }
    return true;
}

bool MapiProfiles::attributeAdd(const char *profile, const char *attribute, QString value) const
{
    if (MAPI_E_SUCCESS != mapi_profile_add_string_attr(m_context, profile, attribute, value.toUtf8())) {
        error() << "profile" << profile << "cannot add attribute" << attribute << value << mapiError();
        return false;
    }
    return true;
}

bool MapiProfiles::attributeModify(const char *profile, const char *attribute, QString value) const
{
    if (MAPI_E_SUCCESS != mapi_profile_modify_string_attr(m_context, profile, attribute, value.toUtf8())) {
        error() << "profile" << profile << "cannot modify attribute" << attribute << value << mapiError();
        return false;
    }
    return true;
}

bool MapiProfiles::attributeRead(struct mapi_profile *profile, const char *attribute, QString &value) const
{
    char **tmp = NULL;
    uint32_t count;

    if (MAPI_E_SUCCESS != GetProfileAttr(profile, attribute, &count, &tmp)) {
        error() << "cannot get attribute" << attribute << mapiError();
        return false;
    }
    if (1 == count) {
        value = QString::fromUtf8(tmp[0]);
    }
    for (uint32_t i = 0; i < count; i++) {
        talloc_free(tmp[i]);
    }
    if (tmp) {
        talloc_free(tmp);
    }
    if (1 == count) {
        return true;
    } else {
        error() << "unexpected attribute count" << attribute << count << mapiError();
        return false;
    }
}

QDebug MapiProfiles::debug() const
{
    static QString prefix = QString::fromAscii("MapiProfiles:");
    return TallocContext::debug(prefix);
}

QString MapiProfiles::defaultGet()
{
    if (!init()) {
        return QString();
    }

    char *profname;
    if (MAPI_E_SUCCESS != GetDefaultProfile(m_context, &profname)) {
        error() << "cannot get default profile:" << mapiError();
        return QString();
    }
    return QString::fromLocal8Bit(profname);
}

bool MapiProfiles::defaultSet(QString profile)
{
    if (!init()) {
        return false;
    }

    if (MAPI_E_SUCCESS != SetDefaultProfile(m_context, profile.toUtf8())) {
        error() << "cannot set default profile:" << profile << mapiError();
        return false;
    }
    return true;
}

QDebug MapiProfiles::error() const
{
    static QString prefix = QString::fromAscii("MapiProfiles:");
    return TallocContext::error(prefix);
}

bool MapiProfiles::init()
{
    if (m_initialised) {
        return true;
    }
    QString profilePath(QDir::home().path());
    profilePath.append(QLatin1String("/.openchange/"));
    
    QString profileFile(profilePath);
    profileFile.append(QString::fromLatin1("profiles.ldb"));

    // Check if the store exists.
    QDir path(profilePath);
    if (!path.exists()) {
        if (path.mkpath(profilePath)) {
            error() << "cannot make profile path:" << profilePath;
            return false;
        }
    }
    if (!QFile::exists(profileFile)) {
        if (MAPI_E_SUCCESS != CreateProfileStore(profileFile.toUtf8(), mapi_profile_get_ldif_path())) {
            error() << "cannot create profile store:" << profileFile << mapiError();
            return false;
        }
    }
    if (MAPI_E_SUCCESS != MAPIInitialize(&m_context, profileFile.toLatin1())) {
        error() << "cannot init profile store:" << profileFile << mapiError();
        return false;
    }
    m_initialised = true;
    return true;
}

QStringList MapiProfiles::list()
{
    if (!init()) {
        return QStringList();
    }

    struct SRowSet proftable;

    if (MAPI_E_SUCCESS != GetProfileTable(m_context, &proftable)) {
        error() << "cannot get profile table" << mapiError();
        return QStringList();
    }

    // debug() << "Profiles in the database:" << proftable.cRows;
    QStringList profiles;
    for (unsigned count = 0; count < proftable.cRows; count++) {
        const char *name = proftable.aRow[count].lpProps[0].value.lpszA;
        uint32_t dflt = proftable.aRow[count].lpProps[1].value.l;

        profiles.append(QString::fromLocal8Bit(name));
        if (dflt) {
            debug() << "default profile:" << name;
        }
    }
    return profiles;
}

bool MapiProfiles::read(const QString &profile, QString &username, QString &domain, QString &server)
{
    if (!init()) {
        return false;
    }

    struct mapi_profile p;
    if (MAPI_E_SUCCESS != OpenProfile(m_context, &p, profile.toUtf8(), NULL)) {
        error() << "cannot open profile:" << profile << mapiError();
        return false;
    }

    if (!attributeRead(&p, "username", username)) {
        return false;
    }

    if (!attributeRead(&p, "domain", domain)) {
        return false;
    }

    if (!attributeRead(&p, "binding", server)) {
        return false;
    }
    return true;
}

bool MapiProfiles::remove(QString profile)
{
    if (!init()) {
        return false;
    }

    if (MAPI_E_SUCCESS != DeleteProfile(m_context, profile.toUtf8())) {
        error() << "cannot delete profile:" << profile << mapiError();
        return false;
    }
    return true;
}

bool MapiProfiles::updatePassword(const QString &profile, const QString &oldPassword, const QString &newPassword)
{
    if (!init()) {
        return false;
    }

    if (MAPI_E_SUCCESS != ChangeProfilePassword(m_context, profile.toUtf8(), oldPassword.toUtf8(), newPassword.toUtf8())) {
        error() << "cannot change password profile:" << profile << mapiError();
        return false;
    }
    return true;
}

bool MapiProfiles::update(const QString &profile, const QString &username, const QString &password, const QString &domain, const QString &server)
{
    if (!init()) {
        return false;
    }
    qDebug() << "Updating profile:"<<profile;
    QString tmpProfile = QLatin1String("_t_m_p_");

    // Save the profile to a temporary name.
    if (MAPI_E_SUCCESS != DuplicateProfile(m_context, profile.toUtf8(), tmpProfile.toUtf8(), username.toUtf8())) {
        error() << "cannot duplicate tmp profile:" << profile << mapiError();
        return false;
    }
    if (!updateInterim(tmpProfile.toUtf8(), username, password, domain, server)) {
        if (!remove(tmpProfile)) {
            return false;
        }
        return false;
    }

    // Rename the temporary modified profile to the original name.
    if (MAPI_E_SUCCESS != DeleteProfile(m_context, profile.toUtf8())) {
        error() << "cannot delete original profile:" << profile << mapiError();
        if (!remove(tmpProfile)) {
            return false;
        }
        return false;
    }
    if (MAPI_E_SUCCESS != RenameProfile(m_context, tmpProfile.toUtf8(), profile.toUtf8())) {
        error() << "cannot rename tmp profile:" << profile << mapiError();
        return false;
    }
    return true;
}

bool MapiProfiles::updateInterim(const char *profile, const QString &username, const QString &password, const QString &domain, const QString &server)
{
    if (!attributeModify(profile, "username", username)) {
        return false;
    }
    if (!attributeModify(profile, "binding", server)) {
        return false;
    }
    if (!attributeModify(profile, "domain", domain)) {
        return false;
    }

    struct mapi_session *session = NULL;
    if (MAPI_E_SUCCESS != MapiLogonProvider(m_context, &session, profile, password.toUtf8(), PROVIDER_ID_NSPI)) {
        error() << "cannot get logon provider" << mapiError();
        return false;
    }

    int retval = ProcessNetworkProfile(session, username.toUtf8().constData(), profileSelectCallback, NULL);
    if (retval != MAPI_E_SUCCESS && retval != 0x1) {
        error() << "cannot process network profile, deleting profile..." << mapiError();
        return false;
    }
    return true;
}

TallocContext::TallocContext(const char *name)
{
    m_ctx = talloc_named(NULL, 0, "%s", name);
    if (!m_ctx) {
        qCritical() << name << "talloc_named failed";
    }
}

TallocContext::~TallocContext()
{
    talloc_free(m_ctx);
}

TALLOC_CTX *TallocContext::ctx()
{
    return m_ctx;
}

QDebug TallocContext::debug(const QString &caller) const
{
    static QString prefix = QString::fromAscii("%1.%2");
    QString talloc = QString::fromAscii(talloc_get_name(m_ctx));
    return qDebug() << prefix.arg(talloc).arg(caller);
}

QDebug TallocContext::error(const QString &caller) const
{
    static QString prefix = QString::fromAscii("%1.%2");
    QString talloc = QString::fromAscii(talloc_get_name(m_ctx));
    return qCritical() << prefix.arg(talloc).arg(caller);
}

char *TallocContext::string(const QString &original)
{
    return talloc_strdup(m_ctx, original.toUtf8().data());
}
