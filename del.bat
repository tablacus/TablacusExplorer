del TE.ncb
attrib -h TE.suo
del TE.suo
del /q TE\TE.vcproj.*.user
del /q Debug\TEd32.*
del /q Debug\TEd64.*
del /q Debug\TE32.*
del /q Debug\TE64.*
del /q TE\Debug\*
rmdir TE\Debug
del /q TE\Release\*
rmdir TE\Release
del /q /s TE\x64\*
rmdir TE\x64\Release
rmdir TE\x64\Debug
rmdir TE\x64
pause
