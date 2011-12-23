project(mapibrowser)

#include_directories(${LIBMAPI_INCLUDE_DIRS} ${QT_INCLUDES} ../connector ${CMAKE_CURRENT_BINARY_DIR})

set(mapibrowser_SRCS mainwindow.cpp mapibrowser.cpp main.cpp ${RESOURCE_EXCHANGE_CONNECTOR_SOURCES})
set(mapibrowser_UI_SRCS mainwindow.ui)

QT4_WRAP_UI(mapibrowser_SRCS ${mapibrowser_UI_SRCS})

qt4_automoc(${mapibrowser_SRCS})
add_executable(mapibrowser ${mapibrowser_SRCS})
target_link_libraries(mapibrowser 
		${LIBMAPI_LIBRARY}
		${KDEPIMLIBS_KPIMUTILS_LIBS}
		${LIBDCERPC_LIBRARY}
		${QT_QTCORE_LIBRARY} 
		${QT_QTGUI_LIBRARY})
