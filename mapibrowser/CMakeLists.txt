project(mapibrowser)

#include_directories(${LIBMAPI_INCLUDE_DIRS} ${QT_INCLUDES} ../connector ${CMAKE_CURRENT_BINARY_DIR})

set(mapibrowser_SRCS 
    mainwindow.cpp 
    mapibrowser.cpp 
    main.cpp
    ${RESOURCE_EXCHANGE_CONNECTOR_SOURCES}
    ${RESOURCE_EXCHANGE_UI_SOURCES}
)
set(mapibrowser_UI_SRCS mainwindow.ui ${RESOURCE_EXCHANGE_UI_FILES})

#QT4_WRAP_UI(mapibrowser_SRCS ${mapibrowser_UI_SRCS})
kde4_add_ui_files( mapibrowser_SRCS ${mapibrowser_UI_SRCS} )

#qt4_automoc(${mapibrowser_SRCS})
#add_executable(mapibrowser ${mapibrowser_SRCS})
kde4_add_executable(mapibrowser RUN_UNINSTALLED ${mapibrowser_SRCS})

target_link_libraries(mapibrowser 
    ${LIBMAPI_LIBRARY}
    ${libmapi_LIBRARIES}
    ${KDEPIMLIBS_KPIMUTILS_LIBS}
    ${LIBDCERPC_LIBRARY}
    ${QT_QTCORE_LIBRARY} 
    ${QT_QTNETWORK_LIBRARY} 
    ${KDE4_KDEUI_LIBS}
    ${KDE4_KDECORE_LIBS})

