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
#define VER_M		12
#define VER_D		1
#else
//Version(DLL)
#define VER_Y		22
#define VER_M		1
#define VER_D		21
#endif

//Icon
#define IDI_TE		1

//Define
//#define USE_TEOBJ
//#define USE_SHELLBROWSER
//#define USE_OBJECTAPI
//#define USE_APIHOOK
//#define USE_HTMLDOC
//#define USE_TESTOBJECT
//#define USE_TESTPATHMATCHSPEC
//#define CHECK_HANDLELEAK
//#define USE_LOG
#define EMULATE_XP	//FALSE &&
#ifndef _WIN64
#define _2000XP
//#define _W2000
#endif
#ifdef _DEBUG
#define _EXEONLY
#endif
