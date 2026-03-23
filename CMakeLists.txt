cmake_minimum_required(VERSION 3.15)

set(API_ROOT "C:/Siemens/Automation/WinCC_OA/3.20/api" CACHE FILEPATH "directory of the WinCC_OA API installation")
include(${API_ROOT}/CMakeDefines.txt)

project(CtrlExcelReader)

set(TARGET CtrlExcelReader)

set(SOURCES
ExternHdl.cxx
)

if ( WIN32 )
  set (SOURCES ${SOURCES} VersInfo.rc)
endif()

add_ctrlext(${TARGET} ${SOURCES})

# --- OpenXLSX (Excel reader / writer) ----------------------------------------
find_package(OpenXLSX CONFIG REQUIRED)
target_link_libraries(${TARGET} OpenXLSX::OpenXLSX)
