#pragma once

#include <grpcpp/grpcpp.h>

#include "dmsoft.pb.h"
#include "dmsoft.grpc.pb.h"

#include "dmObj.h"

class dmsoftWarpper {
public:
    bool isWork = false;
    dmsoft* dm = nullptr;
};

extern std::vector<dmsoftWarpper> dmsofts;
//dmosfts固定大小,需要做成一个池,每次请求一个dmsoft对象,用完后归还,需要定义哪些方法?


std::wstring string2Wstring(const std::string& str);
std::string wstring2String(const std::wstring& wstr);
dmsoft* getDmsoft(int index);

//using namespace dmsoftRpc;



class DmRpcImpl final : public dmsoftRpc::dmsoftService::Service
{
    ::grpc::Status GetDmIndex(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::GetDmIndexResponse* response);
    // 归还dm
    ::grpc::Status ReturnDmIndex(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    ::grpc::Status Ver(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::VerResponse* response);
    // long GetCursorPos(x,y)
    ::grpc::Status GetCursorPos(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::PointResponse* response);
    // long KeyDown(vk_code)
    ::grpc::Status KeyDown(::grpc::ServerContext* context, const ::dmsoftRpc::KeyDownRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long KeyDownChar(key_str)
    ::grpc::Status KeyDownChar(::grpc::ServerContext* context, const ::dmsoftRpc::KeyDownCharRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long KeyPress(vk_code)
    ::grpc::Status KeyPress(::grpc::ServerContext* context, const ::dmsoftRpc::KeyPressRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long KeyUp(vk_code)
    ::grpc::Status KeyUp(::grpc::ServerContext* context, const ::dmsoftRpc::KeyUpRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long LeftClick()
    ::grpc::Status LeftClick(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long LeftDoubleClick()
    ::grpc::Status LeftDoubleClick(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long LeftDown()
    ::grpc::Status LeftDown(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long LeftUp()
    ::grpc::Status LeftUp(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long MoveR(rx,ry)
    ::grpc::Status MoveR(::grpc::ServerContext* context, const ::dmsoftRpc::PointRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long MoveTo(x,y)
    ::grpc::Status MoveTo(::grpc::ServerContext* context, const ::dmsoftRpc::PointRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // string MoveToEx(x,y,w,h)
    ::grpc::Status MoveToEx(::grpc::ServerContext* context, const ::dmsoftRpc::RectRequest* request, ::dmsoftRpc::StringResponse* response);
    // long RightClick()
    ::grpc::Status RightClick(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long RightDown()
    ::grpc::Status RightDown(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long RightUp()
    ::grpc::Status RightUp(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long FindMulColor(x1, y1, x2, y2, color, sim)
    ::grpc::Status FindMulColor(::grpc::ServerContext* context, const ::dmsoftRpc::FindMultiColorRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long FindMultiColor(x1, y1, x2, y2,first_color,offset_color,sim, dir,int32X,int32Y)
    ::grpc::Status FindMultiColor(::grpc::ServerContext* context, const ::dmsoftRpc::FindMultiColorExRequest* request, ::dmsoftRpc::FindMultiColorResponse* response);
    // long FindPic(x1, y1, x2, y2, pic_name, delta_color,sim, dir,int32X, int32Y)
    ::grpc::Status FindPic(::grpc::ServerContext* context, const ::dmsoftRpc::FindPicRequest* request, ::dmsoftRpc::FindPicResponse* response);
    // string GetColor(x,y)
    ::grpc::Status GetColor(::grpc::ServerContext* context, const ::dmsoftRpc::PointRequest* request, ::dmsoftRpc::StringResponse* response);
    // long Reg(reg_code,ver_info)
    ::grpc::Status Reg(::grpc::ServerContext* context, const ::dmsoftRpc::RegRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long SetPath(path)
    ::grpc::Status SetPath(::grpc::ServerContext* context, const ::dmsoftRpc::StringRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long BindWindowEx(hwnd,display,mouse,keypad,public,mode)
    ::grpc::Status BindWindowEx(::grpc::ServerContext* context, const ::dmsoftRpc::BindWindowExRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long DownCpu(type,rate)
    ::grpc::Status DownCpu(::grpc::ServerContext* context, const ::dmsoftRpc::DownCpuRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long EnableBind(enable)
    ::grpc::Status EnableBind(::grpc::ServerContext* context, const ::dmsoftRpc::EnableBindRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long IsBind(hwnd)
    ::grpc::Status IsBind(::grpc::ServerContext* context, const ::dmsoftRpc::IsBindRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long LockInput(lock)
    ::grpc::Status LockInput(::grpc::ServerContext* context, const ::dmsoftRpc::LockRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long UnBindWindow()
    ::grpc::Status UnBindWindow(::grpc::ServerContext* context, const ::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long MoveWindow(hwnd,x,y) 
    ::grpc::Status MoveWindow(::grpc::ServerContext* context, const ::dmsoftRpc::MoveWindowRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long SetClientSize(hwnd,width,height) 
    ::grpc::Status SetClientSize(::grpc::ServerContext* context, const ::dmsoftRpc::SetClientSizeRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long SetWindowSize(hwnd,width,height) 
    ::grpc::Status SetWindowSize(::grpc::ServerContext* context, const ::dmsoftRpc::SetWindowSizeRequest* request, ::dmsoftRpc::IsOkResponse* response);
    // long SetWindowText(hwnd,title) 
    ::grpc::Status SetWindowText(::grpc::ServerContext* context, const ::dmsoftRpc::SetWindowTextRequest* request, ::dmsoftRpc::IsOkResponse* response);
};