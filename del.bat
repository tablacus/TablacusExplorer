del TE2008.ncb
attrib -h TE2008.suo
attrib -h TE*.suo
del TE*.suo
del TE*.sdf
del /q TE\TE.vcproj.*.user
del /q Debug\TEd32.*
del /q Debug\TEd64.*
del /q Debug\TE32.*
del /q Debug\TE64.*
rmdir /s /q TE\Debug
rmdir /s /q TE\Release
rmdir /s /q TE\x64
rmdir /s /q ipch
rmdir /s /q .vs
