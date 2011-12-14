project(exmail)

set( exmailresource_SRCS
  exmailresource.cpp
  ${RESOURCE_EXCHANGE_CONNECTOR_SOURCES}
  ${RESOURCE_EXCHANGE_UI_SOURCES}
)

kde4_add_ui_files( exmailresource_SRCS ${RESOURCE_EXCHANGE_UI_FILES} )
install( FILES exmailresource.desktop DESTINATION "${CMAKE_INSTALL_PREFIX}/share/akonadi/agents" )

kde4_add_kcfg_files(exmailresource_SRCS settings.kcfgc)
kcfg_generate_dbus_interface(${CMAKE_CURRENT_SOURCE_DIR}/exmailresource.kcfg org.kde.Akonadi.ExchangeMail.Settings)
qt4_add_dbus_adaptor(exmailresource_SRCS
  ${CMAKE_CURRENT_BINARY_DIR}/org.kde.Akonadi.ExchangeMail.Settings.xml settings.h Settings
)

kde4_add_executable(akonadi_exmail_resource RUN_UNINSTALLED ${exmailresource_SRCS})

target_link_libraries(akonadi_exmail_resource 
					${KDE4_AKONADI_LIBS} 
					${KDEPIMLIBS_KMIME_LIBS}
					${LIBMAPI_LIBRARY}
					${LIBDCERPC_LIBRARY}
					${QT_QTCORE_LIBRARY} 
					${QT_QTDBUS_LIBRARY} 
					${KDE4_KDECORE_LIBS})

install(TARGETS akonadi_exmail_resource ${INSTALL_TARGETS_DEFAULT_ARGS})
