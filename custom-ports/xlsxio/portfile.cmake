vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO jumoog/xlsxio
    REF ab3bbd4ab642d0d36014ff177a612add048adb56
    SHA512 4b402e90ffdcf38b610e23c7aa335958e486b9bd3aa56b7eb6a207e3ee28eefe791d24423dd7c083ae7555ff78c90a0a40e817e65e7b081e359618d4a9ad5bb9
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
