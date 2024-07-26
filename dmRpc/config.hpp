#pragma once
#include <string>
#include <Windows.h>
#include <filesystem>

using string = std::string;

//使用winapi读写配置文件

//生成默认配置文件
void createDefaultConfigFile();
//读取配置文件
std::tuple<int, string, string, string, string, string, string> readConfigFile();

//WritePrivateProfileStringA默认路径是c:\windows下
auto configFilePath = (std::filesystem::current_path().string() + "dmRpcConfig.ini").c_str();

void createDefaultConfigFile() {

	WritePrivateProfileStringA("server", "ip", "127.0.0.1", configFilePath);
	WritePrivateProfileStringA("server","port", "50051", configFilePath);

	WritePrivateProfileStringA("dm", "DmPath", "c:\\dm7.dl", configFilePath);
	WritePrivateProfileStringA("dm", "regDllPath", "c:\\DmReg.dl", configFilePath);
	WritePrivateProfileStringA("dm", "regCode", "注册码", configFilePath);
	WritePrivateProfileStringA("dm", "regVer", "附加码", configFilePath);
}

std::tuple<int,string, string, string, string, string, string> readConfigFile() {
	if (std::filesystem::exists(configFilePath) == false) {
		createDefaultConfigFile();
	}

	CHAR ip[100] = { 0 };
	CHAR port[100] = { 0 };
	CHAR DmPath[100] = { 0 };
	CHAR regDllPath[100] = { 0 };
	CHAR regCode[100] = { 0 };
	CHAR regVer[100] = { 0 };

	GetPrivateProfileStringA("server", "ip", NULL, ip, 100, configFilePath);
	GetPrivateProfileStringA("server", "port", NULL, port, 100, configFilePath);

	GetPrivateProfileStringA("dm", "DmPath", NULL, DmPath, 100, configFilePath);
	GetPrivateProfileStringA("dm", "regDllPath", NULL, regDllPath, 100, configFilePath);
	GetPrivateProfileStringA("dm", "regCode", NULL, regCode, 100, configFilePath);
	GetPrivateProfileStringA("dm", "regVer", NULL, regVer, 100, configFilePath);

	return { 1, ip, port, DmPath, regDllPath, regCode, regVer };
}