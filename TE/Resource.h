#define PRODUCTNAME				"Tablacus Explorer"

#ifndef IDC_STATIC
#define IDC_STATIC (-1)
#endif

//Version
#define STRING(str) STRING2(str)
#define STRING2(str) #str
#ifdef _EXE
//Version(EXE)
#define VER_Y		21
#define VER_M		11
#define VER_D		1
#else
//Version(DLL)
#define VER_Y		21
#define VER_M		11
#define VER_D		14
#endif

//Icon
#define IDI_TE		1

//Define
//#define _USE_TEOBJ
//#define _USE_SHELLBROWSER
//#define _USE_OBJECTAPI
//#define _USE_APIHOOK
//#define _USE_HTMLDOC
//#define _USE_TESTOBJECT
//#define _USE_TESTPATHMATCHSPEC
//#define _USE_FAKESHELLFOLDER2
//#define _CHECK_HANDLELEAK
//#define _LOG
#define _Emulate_XP_	//FALSE &&
#ifndef _WIN64
#define _2000XP
//#define _W2000
#endif
#ifdef _DEBUG
#define _EXEONLY
#endif
