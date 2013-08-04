del TE.ncb
attrib -h TE.suo
del TE.suo
del /q TE\TE.vcproj.*.user
del /q Debug\TE.*
del /q TE\Debug\*
rmdir TE\Debug
del /q Release\*
rmdir Release
del /q TE\Release\*
rmdir TE\Release
del /q /s x64\Release\*
rmdir x64\Release
del /q /s x64\Debug\*
rmdir x64\Debug
rmdir x64
del /q /s TE\x64\*
rmdir TE\x64\Release
rmdir TE\x64\Debug
rmdir TE\x64
pause
