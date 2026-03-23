vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO jumoog/xlsxio
    REF 62de5a159cfa71291dd5cacf270377d6b6178ecb
    SHA512 fe3daaa8998c1eecbbe246d3188239f3419a392aced982e54063e94108ae99dece618561b50fba3f8eb46ea28ff2f66594d529d718c8d7ebbeea4a20fd6032e9
    HEAD_REF master
    PATCHES
        fix-dependencies.patch
)

file(REMOVE "${SOURCE_PATH}/CMake/FindMinizip.cmake")

string(COMPARE EQUAL "${VCPKG_LIBRARY_LINKAGE}" "static" BUILD_STATIC)
string(COMPARE EQUAL "${VCPKG_LIBRARY_LINKAGE}" "dynamic" BUILD_SHARED)

vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        -DCMAKE_POLICY_DEFAULT_CMP0012=NEW
        -DBUILD_SHARED=${BUILD_SHARED}
        -DBUILD_STATIC=${BUILD_STATIC}
        -DWITH_WIDE=OFF
        -DBUILD_DOCUMENTATION=OFF
        -DBUILD_EXAMPLES=OFF
        -DBUILD_PC_FILES=OFF
        -DBUILD_TOOLS=OFF
)

vcpkg_cmake_install()
vcpkg_copy_pdbs()

vcpkg_cmake_config_fixup(CONFIG_PATH cmake)
vcpkg_fixup_pkgconfig()

file(REMOVE_RECURSE "${CURRENT_PACKAGES_DIR}/debug/include")

vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")
