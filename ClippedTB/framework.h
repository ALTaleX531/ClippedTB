#pragma once

#define WIN32_LEAN_AND_MEAN             // 从 Windows 头文件中排除极少使用的内容
// Windows 头文件
#include <windows.h>
#include <windowsx.h>
#include <comutil.h>
#include <oleacc.h>
#include <shellapi.h>
#include <Shlwapi.h>
#include <ShlObj.h>
#include <taskschd.h>
#include <uxtheme.h>
#include <dwmapi.h>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "Oleacc.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "comctl32.lib")