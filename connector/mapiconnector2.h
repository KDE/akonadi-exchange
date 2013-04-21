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

#ifndef MAPICONNECTOR2_H
#define MAPICONNECTOR2_H

#include <QBitArray>
#include <QDateTime>
#include <QDebug>
#include <QList>
#include <QMap>
#include <QString>

extern "C" {
// libmapi is a C library and must therefore be included that way
// otherwise we'll get linker errors due to C++ name mangling
#include <libmapi/libmapi.h>
}

/**
 * Stringified errors.
 */
extern QString mapiError();

/**
 * Enumerate all starting points in the MAPI store.
 */
typedef enum
{
    MailboxRoot             = olFolderMailboxRoot,          // To navigate the whole tree.
    TopInformationStore     = olFolderTopInformationStore,  // All email note items.
    DeletedItems            = olFolderDeletedItems,
    Outbox                  = olFolderOutbox,
    SentMail                = olFolderSentMail,
    Inbox                   = olFolderInbox,
    CommonView              = olFolderCommonView,
    Calendar                = olFolderCalendar,             // All calendar items.
    Contacts                = olFolderContacts,
    Journal                 = olFolderJournal,
    Notes                   = olFolderNotes,
    Tasks                   = olFolderTasks,
    Drafts                  = olFolderDrafts,
    FoldersAllPublicFolders = olPublicFoldersAllPublicFolders,
    Conflicts               = olFolderConflicts,
    SyncIssues              = olFolderSyncIssues,
    LocalFailures           = olFolderLocalFailures,
    ServerFailures          = olFolderServerFailures,
    Junk                    = olFolderJunk,
    Finder                  = olFolderFinder,
    PublicRoot              = olFolderPublicRoot,
    PublicIPMSubtree        = olFolderPublicIPMSubtree,
    PublicNonIPMSubtree     = olFolderPublicNonIPMSubtree,
    PublicEFormsRoot        = olFolderPublicEFormsRoot,
    PublicFreeBusyRoot      = olFolderPublicFreeBusyRoot,
    PublicOfflineAB         = olFolderPublicOfflineAB,
    PublicEFormsRegistry    = olFolderPublicEFormsRegistry,
    PublicLocalFreeBusy     = olFolderPublicLocalFreeBusy,
    PublicLocalOfflineAB    = olFolderPublicLocalOfflineAB,
    PublicNNTPArticle       = olFolderPublicNNTPArticle,
} MapiDefaultFolder;

/**
 * The id for a MAPI object is (a) hierarchical and (b) associated with a
 * provider. The provider is set with reference to the @ref MapiDefaultFolder
 * and then cascaded down to contained items as we recurse down a folder tree.
 */
class MapiId : public QPair<mapi_id_t, mapi_id_t>
{
public:
    /**
     * For a default folder.
     */
    MapiId(class MapiConnector2 *connection, MapiDefaultFolder folderType);

    /**
     * For a child object.
     */
    MapiId(const MapiId &parent, const mapi_id_t &child);

    /**
     * From an Akonadi id.
     */
    MapiId(const QString &id);

    /**
     * To Akonadi id.
     */
    QString toString() const;

    bool isValid() const;

    static const QChar fidIdSeparator;

private:
    enum Provider
    {
        INVALID,
        EMSDB,
        NSPI
    } m_provider;
    friend class MapiConnector2;
};

/**
 * A class which wraps a talloc memory allocator such that objects of this type
 * automatically free the used memory on destruction.
 */
class TallocContext
{
public:
    TallocContext(const char *name);

    virtual ~TallocContext();

    TALLOC_CTX *ctx();

    /**
     * Make a talloc-friendly allocation.
     */
    template <class T>
    T *allocate();

    /**
     * Make a talloc-friendly array.
     */
    template <class T>
    T *array(unsigned size);

    /**
     * Make a talloc-friendly UTF-8 encoded string.
     */
    char *string(const QString &original);

protected:
    TALLOC_CTX *m_ctx;

    /**
     * Debug and error reporting. Each subclass should reimplement with 
     * logic that emits a prefix identifying the object involved. 
     */
    virtual QDebug debug() const = 0;
    virtual QDebug error() const = 0;

    /**
     * A couple of helper functions that subclasses can use to implement the
     * above virtual methods.
     */
    QDebug debug(const QString &caller) const;
    QDebug error(const QString &caller) const;
};

template <class T>
T *TallocContext::allocate()
{
    return talloc(m_ctx, T);
}

template <class T>
T *TallocContext::array(unsigned size)
{
    return talloc_array(m_ctx, T, size);
}

/**
 * A class for managing access to MAPI profiles. This is the root for all other
 * MAPI interactions.
 */
class MapiProfiles : protected TallocContext
{
public:
    MapiProfiles();
    virtual ~MapiProfiles();

    /**
     * Find existing profiles.
     */
    QStringList list();

    /**
     * Set the default profile.
     */
    bool defaultSet(QString profile);

    /**
     * Get the default profile.
     */
    QString defaultGet();

    /**
     * Add a new profile.
     */
    bool add(const QString &profile, const QString &username, const QString &password, const QString &domain, const QString &server);

    bool read(const QString &profile, QString &username, QString &domain, QString &server);

    bool update(const QString &profile, const QString &username, const QString &password, const QString &domain, const QString &server);

    /**
     * Modify an existing profile's password.
     */
    bool updatePassword(const QString &profile, const QString &oldPassword, const QString &newPassword);

    /**
     * Remove a profile.
     */
    bool remove(QString profile);

protected:
    mapi_context *m_context;

    /**
     * Must be called first!
     */
    bool init();

private:
    bool m_initialised;

    bool attributeAdd(const char *profile, const char *attribute, QString value) const;
    bool attributeModify(const char *profile, const char *attribute, QString value) const;
    bool attributeRead(struct mapi_profile *profile, const char *attribute, QString &value) const;
    bool updateInterim(const char *profile, const QString &username, const QString &password, const QString &domain, const QString &server);

    virtual QDebug debug() const;
    virtual QDebug error() const;
};

/**
 * The main class represents a connection to the MAPI server.
 */
class MapiConnector2 : protected QObject, public MapiProfiles
{
    Q_OBJECT
public:
    MapiConnector2();
    virtual ~MapiConnector2();

    /**
     * Connect to the server.
     * 
     * @param profile   Name in libmapi database.
     * @return          True if the connection attempt succeeds.
     */
    bool login(QString profile);

    /**
     * Factory for getting default folder ids. MAPI has two kinds of folder,
     * Public and user-specific. This wraps the two together.
     */
    bool defaultFolder(MapiDefaultFolder folderType, MapiId *id);

    /**
     * How many entries are there in the GAL?
     */
    bool GALCount(unsigned *totalCount);

    /**
     * Fetch upto the requested number of entries from the GAL. The start
     * point is where we previously left off.
     */
    bool GALRead(unsigned requestedCount, SPropTagArray *tags, SRowSet **results, unsigned *percentagePosition = 0);

    bool GALSeek(const QString &displayName, unsigned *percentagePosition = 0, SPropTagArray *tags = 0, SRowSet **results = 0);

    bool GALRewind();

    mapi_object_t *store(const MapiId &id)
    {
        switch (id.m_provider)
        {
        case MapiId::EMSDB:
            return m_store;
        case MapiId::NSPI:
            return m_nspiStore;
        default:
            return 0;
        }
    }

    /**
     * Resolve the given names.
     * 
     * @param names A 0-terminated array of 0-terminated names to be resolved.
     * @param tags  The item properties we are interested in.
     */
    bool resolveNames(const char *names[], SPropTagArray *tags,
              SRowSet **results, PropertyTagArray_r **statuses);

private:
    mapi_object_t openFolder(mapi_id_t folderID);

    mapi_session *m_session;
    mapi_object_t *m_store;
    mapi_object_t *m_nspiStore;
    class QSocketNotifier *m_notifier;

    virtual QDebug debug() const;
    virtual QDebug error() const;

private Q_SLOTS:
    void notified(int fd);
};

#endif // MAPICONNECTOR2_H
