cmake_minimum_required(VERSION 3.20)
project(7ZipContextMenu LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Create DLL
add_library(7ZipContextMenu SHARED
    src/ContextMenu.cpp
    src/ContextMenu.h
    src/ContextMenu.def
)

target_compile_definitions(7ZipContextMenu PRIVATE
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
)

# Ensure proper UTF-8 handling for source files
if(MSVC)
    target_compile_options(7ZipContextMenu PRIVATE /utf-8)
endif()

target_link_libraries(7ZipContextMenu PRIVATE
    shlwapi
    pathcch
    ole32
    shell32
)

# Install configuration - install to build/install
set(CMAKE_INSTALL_PREFIX "${CMAKE_BINARY_DIR}/install" CACHE PATH "Install path" FORCE)

# Install DLL and PDB
install(TARGETS 7ZipContextMenu
    RUNTIME DESTINATION .
    LIBRARY DESTINATION .
)
install(FILES $<TARGET_PDB_FILE:7ZipContextMenu>
    DESTINATION .
    OPTIONAL
)

# Install scripts
install(FILES
    scripts/install.ps1
    scripts/uninstall.ps1
    DESTINATION .
)

# Install package assets
install(FILES Package/AppxManifest.xml DESTINATION .)
install(DIRECTORY Package/Assets DESTINATION .)
