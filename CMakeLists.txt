cmake_minimum_required(VERSION 3.20)
project(longpad LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(longpad WIN32 main.cpp)

target_link_libraries(longpad
    comctl32
    comdlg32
    shell32
    advapi32
)

# Release: optimize for size/speed, no console window
target_compile_options(longpad PRIVATE
    $<$<CONFIG:Release>:/O2 /GL>
)
target_link_options(longpad PRIVATE
    $<$<CONFIG:Release>:/LTCG>
)
