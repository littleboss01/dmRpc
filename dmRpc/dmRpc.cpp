// dmRpc.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>

#include <atlbase.h>  // CComBSTR
#include <comutil.h>  // COleVariant
#include <string>
#include <iostream>


//引入grpc?
#include <grpcpp/grpcpp.h>

#include "dmsoft.pb.h"
#include "dmsoft.grpc.pb.h"
#include "DmRpcImpl.h"

#include "dmObjHelp.h"
#include "config.hpp"

using namespace dmsoftRpc;


//大漠调用RegDll.dll实现免注册


//定义rpc客户端
class DmRpcClient
{
	public:
	DmRpcClient(std::shared_ptr<grpc::Channel> channel) : stub_(dmsoftService::NewStub(channel)) {}
	std::string Ver()
	{
		IndexRequest request;
		VerResponse response;
		grpc::ClientContext context;
		request.set_index(0);
		grpc::Status status = stub_->Ver(&context, request, &response);
		if (status.ok())
		{
			return response.version();
		}
		else
		{
			return "RPC failed";
		}
	}
	int GetDmIndex()
	{
		IndexRequest request;
		GetDmIndexResponse response;
		grpc::ClientContext context;
		request.set_index(0);
		grpc::Status status = stub_->GetDmIndex(&context, request, &response);
		if (status.ok())
		{
			return response.index();
		}
		else
		{
			return -1;
		}
	}
	int ReturnDmIndex(int index) {
		IndexRequest request;
		IsOkResponse response;
		grpc::ClientContext context;
		request.set_index(index);
		grpc::Status status = stub_->ReturnDmIndex(&context, request, &response);
		if (status.ok())
		{
			return response.isok();
		}
		else
		{
			return -1;
		}
	}
	private:
		std::unique_ptr<dmsoftService::Stub> stub_;
};
//创建99个大漠对象
//dmsofts = std::vector<dmsoft>(99, dmsoft());
//CoInitializeEx(NULL, 0);
//dmsofts.resize(99);
//for (int i=0;i<dmsofts.size(); i++)
//{
//	dmsofts[i].dm =new dmsoft();
//	dmsofts[i].dm->SetPath( std::to_wstring(i) );
//}
//for (int i = 0; i < dmsofts.size(); i++)
//{
//	std::wcout << dmsofts[i].dm->GetPath() << std::endl;
//}
#include <csignal>
std::unique_ptr<grpc::Server> server = nullptr;

void signalHandler(int signum) {
	if (signum == SIGINT) {
		server->Shutdown();
	}
}

int main()
{
	//读取配置文件
	auto [isok, ip, port, DmPath, regDllPath, regCode, regVer]= readConfigFile();


	std::cout << "DmReg.dll和dm5.dll放到c盘根目录" << std::endl;
	std::cout << "release编译的时候需要在连接器-输入-依赖,填写,Crypt32.lib,静态使用mfc和c++-代码生成-运行库 设置mt" << std::endl;
	dmObjHelp::regDll(DmPath);

	auto dm1 = dmsoft();
	std::wcout << dm1.Ver() << std::endl;

	std::thread t1([&]() {
		std::string server_address( ip + string( ":") + port);
		DmRpcImpl service;
		grpc::ServerBuilder builder;
		builder.AddListeningPort(server_address, grpc::InsecureServerCredentials());
		builder.RegisterService(&service);
		server=builder.BuildAndStart();
		std::cout << "Server listening on " << server_address << std::endl;
		signal(SIGINT, signalHandler);
		server->Wait();
	 
		});
	t1.join();
	
	if (0) {
		t1.detach();
		//启动rpc客户端
		std::string server_address(ip + string(":") + port);
		DmRpcClient client(grpc::CreateChannel(server_address, grpc::InsecureChannelCredentials()));
		std::cout << client.GetDmIndex() << std::endl;
		client.ReturnDmIndex(0);
		std::cout << client.GetDmIndex() << std::endl;
		client.ReturnDmIndex(0);
		std::string version = client.Ver();
		std::cout << "Dmsoft version: " << version << std::endl;
	}

}

// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件


//bool VerifyConversion(const std::string& originalString) {
//	// 将 std::string 转换为 CComBSTR
//	CComBSTR bstrPwd(originalString.c_str());
//
//	// 将 CComBSTR 转换为 COleVariant
//	COleVariant variantPwd(bstrPwd);
//
//	// 确认 COleVariant 包含的是 BSTR 类型的数据
//	if (variantPwd.vt != VT_BSTR) {
//		return false;
//	}
//
//	// 从 COleVariant 中提取 BSTR
//	BSTR bstrFromVariant = variantPwd.bstrVal;
//
//	// 将 BSTR 转换回 std::string
//	_bstr_t bstrWrapper(bstrFromVariant);
//	std::string convertedString((const char*)bstrWrapper);
//
//	// 比较原始字符串和转换后的字符串
//	return originalString == convertedString;
//}