project(akonadi_exchange)
cmake_minimum_required(VERSION 2.8.12)

add_definitions( -DQT_NO_CAST_FROM_ASCII )
add_definitions( -DQT_NO_CAST_TO_ASCII )

# search packages used by KDE
find_package (KDE4 REQUIRED)
include(KDE4Defaults)
find_package (KdepimLibs REQUIRED)


# search for the OpenChange MAPI libraries
set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} ${CMAKE_CURRENT_SOURCE_DIR}/cmake/modules/ )

find_package(LibMapi)
set_package_properties(LibMapi PROPERTIES
    DESCRIPTION "OpenChange MAPI Client Library"
    URL "http://www.openchange.org"
    TYPE REQUIRED
    PURPOSE "Needed to support OpenChange / Exchange servers"
)

find_package(Samba)
set_package_properties(Samba PROPERTIES
    DESCRIPTION "The smbclient library"
    URL "http://www.samba.org"
    TYPE REQUIRED
    PURPOSE "Exchange needs the smbclient library"
)

find_package(Xsltproc)
set_package_properties(Xsltproc PROPERTIES
    DESCRIPTION "XSLT processor from libxslt"
    TYPE REQUIRED
    PURPOSE "Required to generate D-Bus interfaces."
)

find_package(Boost)
set_package_properties(Boost PROPERTIES
    DESCRIPTION "Boost library"
    TYPE REQUIRED
    PURPOSE "Required by akonadi."
)

if(WIN32)
    set(LIB_INSTALL_DIR ${LIB_INSTALL_DIR}
                        RUNTIME DESTINATION ${BIN_INSTALL_DIR}
                        LIBRARY DESTINATION ${LIB_INSTALL_DIR}
                        ARCHIVE DESTINATION ${LIB_INSTALL_DIR})
endif()

set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} ${KDE4_ENABLE_EXCEPTIONS}")


include_directories(
    ${KDE4_INCLUDES}
    ${KDEPIMLIBS_INCLUDE_DIRS}
    ${LibMapi_INCLUDE_DIRS}
    ${LibDcerpc_INCLUDE_DIRS}
    connector
    ui
)

# define global path to the connector sources for every resource to use
set( RESOURCE_EXCHANGE_CONNECTOR_SOURCES
     ${CMAKE_CURRENT_SOURCE_DIR}/connector/mapiconnector2.cpp
     ${CMAKE_CURRENT_SOURCE_DIR}/connector/mapiobjects.cpp
)
# define global path to the UI sources for every resource to use
set( RESOURCE_EXCHANGE_UI_SOURCES
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/profiledialog.cpp
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/createprofiledialog.cpp
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/modifyprofiledialog.cpp
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/changepassworddialog.cpp
)
# define global path to the UI files for every resource to use
set( RESOURCE_EXCHANGE_UI_FILES
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/profiledialog.ui
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/createprofiledialog.ui
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/modifyprofiledialog.ui
     ${CMAKE_CURRENT_SOURCE_DIR}/ui/changepassworddialog.ui
)

add_subdirectory(calendar)
add_subdirectory(contacts)
add_subdirectory(mail)
add_subdirectory(mapibrowser)

feature_summary(WHAT ALL
                     INCLUDE_QUIET_PACKAGES
                     FATAL_ON_MISSING_REQUIRED_PACKAGES
               )
