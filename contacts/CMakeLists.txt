project(exgal)

set( exgalresource_SRCS
    exgalresource.cpp
    ${RESOURCE_EXCHANGE_CONNECTOR_SOURCES}
    ${RESOURCE_EXCHANGE_UI_SOURCES}
    ${CMAKE_CURRENT_SOURCE_DIR}/../connector/mapiresource.cpp
)

kde4_add_ui_files( exgalresource_SRCS ${RESOURCE_EXCHANGE_UI_FILES} )
install( FILES exgalresource.desktop DESTINATION "${CMAKE_INSTALL_PREFIX}/share/akonadi/agents" )

kde4_add_kcfg_files(exgalresource_SRCS settings.kcfgc)
kcfg_generate_dbus_interface(${CMAKE_CURRENT_SOURCE_DIR}/exgalresource.kcfg org.kde.Akonadi.ExGal.Settings)
qt4_add_dbus_adaptor(exgalresource_SRCS
    ${CMAKE_CURRENT_BINARY_DIR}/org.kde.Akonadi.ExGal.Settings.xml settings.h Settings
)

kde4_add_executable(akonadi_exgal_resource RUN_UNINSTALLED ${exgalresource_SRCS})

target_link_libraries(akonadi_exgal_resource 
    ${KDEPIMLIBS_AKONADI_LIBS}
    ${KDEPIMLIBS_KABC_LIBS}
    ${KDEPIMLIBS_KMIME_LIBS}
    ${KDEPIMLIBS_KPIMUTILS_LIBS}
    ${LibMapi_LIBRARIES}
    ${LibDcerpc_LIBRARIES}
    ${QT_QTCORE_LIBRARY}
    ${QT_QTDBUS_LIBRARY}
    ${QT_QTNETWORK_LIBRARY}
    ${KDE4_KDECORE_LIBS}
)

install(TARGETS akonadi_exgal_resource ${INSTALL_TARGETS_DEFAULT_ARGS})
