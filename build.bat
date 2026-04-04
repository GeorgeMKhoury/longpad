@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\Common7\Tools\VsDevCmd.bat" -arch=x64
if errorlevel 1 (echo Failed to init VS environment & exit /b 1)
if not exist build mkdir build
cd build
cmake .. -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
if errorlevel 1 (echo CMake configure failed & exit /b 1)
nmake
if errorlevel 1 (echo Build failed & exit /b 1)
echo Build succeeded: build\longpad.exe
