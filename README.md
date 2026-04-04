# Longpad

This is a simple text editor, about as simple as it could possibly be. Think
of it as a notepad clone that can handle long files. No fancy syntax 
highlighting, just the ability to edit large text files quickly.

# Build instructions

Run `build.bat` from a command prompt or powershell in the project directory. 
That script assumes you have Visual Studio Build Tools 2026 installed. 
If you have a different version,
open your Developer PowerShell for VS and run:

```
mkdir build
cd build
cmake .. -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Release
nmake
```
