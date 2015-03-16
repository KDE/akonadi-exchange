# - Try to find the OpenChange MAPI library
# Once done this will define
#
#  LibMapi_FOUND - system has OpenChange MAPI library (libmapi)
#  LibMapi_INCLUDE_DIRS - the libmapi include directories
#  LibMapi_LIBRARIES - Required libmapi link libraries
#  LibMapipp_LIBRARIES - Required libmapi++ link libraries
#  LibMapi_DEFINITIONS - Compiler switches for libmapi
#  LibDcerpc_LIBRARIES - Required libdcerp libraries
#
# Copyright (C) 2007 Brad Hards (bradh@frogmouth.net)
# Copyright © 2015 Christophe Giboudeaux (cgiboudeaux@gmx.com)
#
# Redistribution and use is allowed according to the terms of the BSD license.
# For details see the COPYING-CMAKE-SCRIPTS file in kdelibs/cmake/modules/

find_package(PkgConfig QUIET)
pkg_check_modules(PC_LibMapi QUIET libmapi)
pkg_check_modules(PC_LibMapipp QUIET libmapi++)
pkg_check_modules(PC_LibDcerpc QUIET dcerpc)

find_library(LibMapi_LIBRARIES
    NAMES mapi
    HINTS ${PC_LibMapi_LIBRARY_DIRS}
)

find_library(LibMapipp_LIBRARIES
    NAMES mapipp
    HINTS ${PC_LibMapipp_LIBRARY_DIRS}
)

find_library(LibDcerpc_LIBRARIES
    NAMES dcerpc
    HINTS ${PC_LibDcerpc_LIBRARY_DIRS}
)

find_path(LibMapi_INCLUDE_DIRS
    NAMES libmapi/version.h
    HINTS ${PC_LibMapi_INCLUDE_DIRS}
)

find_path(LibDcerpc_INCLUDE_DIRS
    NAMES dcerpc.h
    HINTS ${PC_LibDcerpc_INCLUDE_DIRS}
          samba-4.0
)

set(LibMapi_DEFINITIONS "${PC_LibMapi_DEFINITIONS}")

# The libmapi.pc and libmapi/version.h don't return the same version,
# we cannot use the pkgconfig one.
# set(LibMapi_VERSION "${PC_LibMapi_VERSION}")

if(EXISTS ${LibMapi_INCLUDE_DIRS}/libmapi/version.h)
  file(READ ${LibMapi_INCLUDE_DIRS}/libmapi/version.h VERSION_H_CONTENT)
  string(REGEX MATCH "#define OPENCHANGE_VERSION_OFFICIAL_STRING[ ]+\"[0-9.]+\"" LIBMAPI_VERSION_STRING ${VERSION_H_CONTENT})
  string(REGEX REPLACE "^.*OPENCHANGE_VERSION_OFFICIAL_STRING[ ]+\"(.*)\"$" "\\1" LibMapi_VERSION ${LIBMAPI_VERSION_STRING})
endif()

include(FindPackageHandleStandardArgs)

find_package_handle_standard_args(LibMapi
    FOUND_VAR LibMapi_FOUND
    REQUIRED_VARS LibMapi_LIBRARIES LibMapipp_LIBRARIES LibDcerpc_LIBRARIES LibDcerpc_INCLUDE_DIRS
    VERSION_VAR LibMapi_VERSION
)

mark_as_advanced(
    LibMapi_LIBRARIES
    LibMapipp_LIBRARIES
    LibMapi_INCLUDE_DIRS
    LibDcerpc_LIBRARIES
    LibDcerpc_INCLUDE_DIRS
    LibMapi_VERSION
    LibMapi_DEFINITIONS
)
