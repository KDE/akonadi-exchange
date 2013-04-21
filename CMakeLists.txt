project(akonadi_exchange)
cmake_minimum_required(VERSION 2.6)

add_definitions( -DQT_NO_CAST_FROM_ASCII )
add_definitions( -DQT_NO_CAST_TO_ASCII )

# search packages used by KDE
find_package (KDE4 REQUIRED)
include(KDE4Defaults)
include(MacroLibrary)
include(MacroOptionalAddSubdirectory)
include(CheckIncludeFiles)
find_package (KdepimLibs REQUIRED)


# search for the OpenChange MAPI libraries
set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} ${CMAKE_CURRENT_SOURCE_DIR}/cmake/modules/ )

macro_optional_find_package(OpenChange)
macro_log_feature(libmapi_FOUND "libmapi" "OpenChange MAPI Client Library" "http://www.openchange.org" FALSE "" "Needed to support OpenChange / Exchange servers")

macro_optional_find_package(Samba)
macro_log_feature(SAMBA_FOUND "Samba" "Samba smbclient" "")

if (NOT libmapi_FOUND)
    message(WARNING "Exchange needs OpenChange MAPI Client Library.")
    return()
endif (NOT libmapi_FOUND)

if (NOT SAMBA_FOUND)
    message(WARNING "Exchange needs Samba smbclient.")
    return()
endif (NOT SAMBA_FOUND)

find_program(XSLTPROC_EXECUTABLE xsltproc)
macro_log_feature(XSLTPROC_EXECUTABLE "xsltproc" "The command line XSLT processor from libxslt" "http://xmlsoft.org/XSLT/" FALSE "" "Needed for building Akonadi resources. Recommended.")

if (XSLTPROC_EXECUTABLE)
  # generates a D-Bus interface description from a KConfigXT file
  macro( kcfg_generate_dbus_interface _kcfg _name )
    add_custom_command(
      OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/${_name}.xml
      COMMAND ${XSLTPROC_EXECUTABLE} --stringparam interfaceName ${_name}
              ${KDEPIMLIBS_DATA_DIR}/akonadi-kde/kcfg2dbus.xsl
              ${_kcfg}
              > ${CMAKE_CURRENT_BINARY_DIR}/${_name}.xml
      DEPENDS ${KDEPIMLIBS_DATA_DIR}/akonadi-kde/kcfg2dbus.xsl
              ${_kcfg}
    )
  endmacro( kcfg_generate_dbus_interface )
endif (XSLTPROC_EXECUTABLE)

if(WIN32)
    set(LIB_INSTALL_DIR ${LIB_INSTALL_DIR}
                        RUNTIME DESTINATION ${BIN_INSTALL_DIR}
                        LIBRARY DESTINATION ${LIB_INSTALL_DIR}
                        ARCHIVE DESTINATION ${LIB_INSTALL_DIR} )
endif(WIN32)


set( CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} ${KDE4_ENABLE_EXCEPTIONS}" )


include_directories(
    connector
    ui
    ${LIBMAPI_INCLUDE_DIRS}
    ${SAMBA_INCLUDE_DIR}/samba-4.0
    ${KDE4_INCLUDES}
    ${KDEPIMLIBS_INCLUDE_DIRS}
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
