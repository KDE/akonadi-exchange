project(excal)

set( excalresource_SRCS
  excalresource.cpp
  ${RESOURCE_EXCHANGE_CONNECTOR_SOURCES}
  ${RESOURCE_EXCHANGE_UI_SOURCES}
  ${CMAKE_CURRENT_SOURCE_DIR}/../connector/mapiresource.cpp
)

kde4_add_ui_files( excalresource_SRCS ${RESOURCE_EXCHANGE_UI_FILES} )
install( FILES excalresource.desktop DESTINATION "${CMAKE_INSTALL_PREFIX}/share/akonadi/agents" )

kde4_add_kcfg_files(excalresource_SRCS settings.kcfgc)
kcfg_generate_dbus_interface(${CMAKE_CURRENT_SOURCE_DIR}/excalresource.kcfg org.kde.Akonadi.ExCal.Settings)
qt4_add_dbus_adaptor(excalresource_SRCS
  ${CMAKE_CURRENT_BINARY_DIR}/org.kde.Akonadi.ExCal.Settings.xml settings.h Settings
)

kde4_add_executable(akonadi_excal_resource RUN_UNINSTALLED ${excalresource_SRCS})

target_link_libraries(akonadi_excal_resource 
					${KDE4_AKONADI_LIBS} 
					${KDEPIMLIBS_KCAL_LIBS}
					kcalutils
					${KDEPIMLIBS_KMIME_LIBS}
					${KDEPIMLIBS_KPIMUTILS_LIBS}
					${LIBMAPI_LIBRARY}
					${libmapi_LIBRARIES}
					${LIBDCERPC_LIBRARY}
					${QT_QTCORE_LIBRARY} 
					${QT_QTDBUS_LIBRARY} 
					${KDE4_KDECORE_LIBS})

install(TARGETS akonadi_excal_resource ${INSTALL_TARGETS_DEFAULT_ARGS})
