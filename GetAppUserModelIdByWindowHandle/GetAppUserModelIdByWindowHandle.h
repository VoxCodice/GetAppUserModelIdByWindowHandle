#pragma once

#include <windows.h>
#include <initguid.h>

extern "C" __declspec(dllexport) bool GetAppUserModelIdByWindowHandle(HWND hWnd, WCHAR* appUserModelId, int appUserModelIdLength);
