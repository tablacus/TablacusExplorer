#define PRODUCTNAME				"Tablacus Explorer"

#ifndef IDC_STATIC
#define IDC_STATIC (-1)
#endif

#define IDI_TE		128

//Version
#define STRING(str) STRING2(str)
#define STRING2(str) #str
#define VER_Y		17
#define VER_M		10
#define VER_D		10

//Define
//#define _USE_BSEARCHAPI
//#define _USE_APIHOOK
//#define _USE_HTMLDOC
//#define _USE_TESTOBJECT
//#define _USE_TESTPATHMATCHSPEC
//#define _LOG
#define _Emulate_XP_	//FALSE &&
#ifndef _WIN64
#define _2000XP
//#define _W2000
#endif
//bug fix 1 for Windows 10 Insider Preview 14986-
#define _FIXWIN10IPBUG1
