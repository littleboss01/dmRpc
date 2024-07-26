#include "dmObjHelp.h"

#include <fstream>
#include <filesystem>
#include <Windows.h>

int dmObjHelp::regDll(string path)
{
	HMODULE hModule= NULL;
	//在当前目录下找不到DmReg.dll,就去c盘根目录找
	auto arrPath={"DmReg.dll","c:\\DmReg.dll"};
	for (auto& item : arrPath)
	{
		if (std::filesystem::exists(item)) {
			hModule = LoadLibraryA(item);
		}
	}
	if (hModule == NULL)
	{
		return -1;
	}
	typedef int(__stdcall* SetDllPathA)(const char*, int);
	typedef int(__stdcall* SetDllPathW)(const TCHAR*, int);
	SetDllPathA setDllPathA = (SetDllPathA)GetProcAddress(hModule, "SetDllPathA");
	if (setDllPathA == NULL)
	{
		return -2;
	}
	int ret = setDllPathA(path.c_str(), 0);
	FreeLibrary(hModule);
	return ret;
}
