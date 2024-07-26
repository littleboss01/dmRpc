#include "DmRpcImpl.h"

//创建99个大漠对象,因为没有注册大漠对象,所以new出来的对象是无效的
//std::vector<dmsoft*> dmsofts = std::vector<dmsoft*>(99, new dmsoft());
std::vector<dmsoftWarpper> dmsofts(99);

// 将 std::string 转换为 std::wstring
std::wstring string2Wstring(const std::string& str) {
    int len = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (len == 0) {
        // 处理错误
        return L"";
    }
    std::wstring wstr(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &wstr[0], len);
    return wstr;
}

// 将 std::wstring 转换为 std::string
std::string wstring2String(const std::wstring& wstr) {
    int len = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (len == 0) {
        // 处理错误
        return "";
    }
    std::string str(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &str[0], len, NULL, NULL);
    return str;
}

dmsoft* getDmsoft(int index)
{
	if (index < 0 || index >= dmsofts.size())
	{
		return nullptr;
	}
	return dmsofts[index].dm;
}

::grpc::Status DmRpcImpl::GetDmIndex(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::GetDmIndexResponse* response)
{
    for (int i = 0; i < dmsofts.size(); i++)
    {
        if (!dmsofts[i].isWork)
        {
            response->set_index(i);
            dmsofts[i].isWork = true;
            dmsofts[i].dm = new dmsoft();
            break;
        }
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::ReturnDmIndex(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    if (request->index() >= 0 && request->index() < dmsofts.size())
    {
        dmsofts[request->index()].dm->UnBindWindow();
        dmsofts[request->index()].dm->FreePic(L"*.bmp");
        dmsofts[request->index()].isWork = false;
        delete dmsofts[request->index()].dm;
        dmsofts[request->index()].dm = nullptr;
        response->set_isok(true);
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::Ver(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::VerResponse* response)
{
    dmsoft* dm = getDmsoft(request->index());
    if (dm) {
        response->set_version(wstring2String(dm->Ver()));
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::GetCursorPos(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::PointResponse* response)
{
    long x=0, y=0;
    dmsoft* dm = getDmsoft(request->index());
    if (dm) {
        dm->GetCursorPos();
        response->set_x(x);
        response->set_y(y);
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::KeyDown(::grpc::ServerContext* context, const::dmsoftRpc::KeyDownRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm= getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->KeyDown(request->vk_code()));
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::KeyDownChar(::grpc::ServerContext* context, const::dmsoftRpc::KeyDownCharRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->KeyDownChar(string2Wstring(request->key_str())));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::KeyPress(::grpc::ServerContext* context, const::dmsoftRpc::KeyPressRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->KeyPress(request->vk_code()));
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::KeyUp(::grpc::ServerContext* context, const::dmsoftRpc::KeyUpRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->KeyUp(request->vk_code()));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::LeftClick(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->LeftClick());
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::LeftDoubleClick(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->LeftDoubleClick());
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::LeftDown(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->LeftDown());
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::LeftUp(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->LeftUp());
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::MoveR(::grpc::ServerContext* context, const::dmsoftRpc::PointRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->MoveR(request->x(), request->y()));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::MoveTo(::grpc::ServerContext* context, const::dmsoftRpc::PointRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->MoveTo(request->x(), request->y()));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::MoveToEx(::grpc::ServerContext* context, const::dmsoftRpc::RectRequest* request, ::dmsoftRpc::StringResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_str(wstring2String(dm->MoveToEx(request->x(), request->y(), request->w(), request->h())));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::RightClick(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->RightClick());
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::RightDown(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->RightDown());
		}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::RightUp(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->RightUp());
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::FindMulColor(::grpc::ServerContext* context, const::dmsoftRpc::FindMultiColorRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->FindMulColor(request->x1(), request->y1(), request->x2(), request->y2(),string2Wstring( request->color().c_str()), request->sim()));
		}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::FindMultiColor(::grpc::ServerContext* context, const::dmsoftRpc::FindMultiColorExRequest* request, ::dmsoftRpc::FindMultiColorResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
       auto [result,x,y] = dm->FindMultiColor(request->x1(), request->y1(), request->x2(), request->y2(), string2Wstring(request->first_color()), string2Wstring(request->offset_color()), request->sim(), request->dir());
       response->set_x(x);
       response->set_y(y);
        response->set_isok(result);
    }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::FindPic(::grpc::ServerContext* context, const::dmsoftRpc::FindPicRequest* request, ::dmsoftRpc::FindPicResponse* response)
{
    long x=-1, y=-1;
    auto dm = getDmsoft(request->index());
    if (dm) {
        auto [isok, x, y] = dm->FindPic(request->x1(), request->y1(), request->x2(), request->y2(), string2Wstring(request->pic_name()), string2Wstring(request->delta_color()), request->sim(), request->dir());
        response->set_x(x);
        response->set_y(y);
		response->set_isok(isok);
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::GetColor(::grpc::ServerContext* context, const::dmsoftRpc::PointRequest* request, ::dmsoftRpc::StringResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_str(wstring2String(dm->GetColor(request->x(), request->y())));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::Reg(::grpc::ServerContext* context, const::dmsoftRpc::RegRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->Reg(string2Wstring(request->reg_code()),string2Wstring(request->ver_info())));
	}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::SetPath(::grpc::ServerContext* context, const::dmsoftRpc::StringRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->SetPath(string2Wstring(request->path())));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::BindWindowEx(::grpc::ServerContext* context, const::dmsoftRpc::BindWindowExRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->BindWindowEx(request->hwnd(), string2Wstring(request->display()),string2Wstring(request->mouse()), string2Wstring(request->keypad()),string2Wstring(request->public_()), request->mode()));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::DownCpu(::grpc::ServerContext* context, const::dmsoftRpc::DownCpuRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		//response->set_isok(dm->DownCpu(request->type(), request->rate())); //TODO 版本不同
		}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::EnableBind(::grpc::ServerContext* context, const::dmsoftRpc::EnableBindRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->EnableBind(request->enable()));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::IsBind(::grpc::ServerContext* context, const::dmsoftRpc::IsBindRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->IsBind(request->hwnd()));
		}
        
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::LockInput(::grpc::ServerContext* context, const::dmsoftRpc::LockRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->LockInput(request->lock()));
		}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::UnBindWindow(::grpc::ServerContext* context, const::dmsoftRpc::IndexRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->UnBindWindow());
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::MoveWindow(::grpc::ServerContext* context, const::dmsoftRpc::MoveWindowRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->MoveWindow(request->hwnd(), request->x(), request->y()));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::SetClientSize(::grpc::ServerContext* context, const::dmsoftRpc::SetClientSizeRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
		response->set_isok(dm->SetClientSize(request->hwnd(), request->width(), request->height()));
		}
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::SetWindowSize(::grpc::ServerContext* context, const::dmsoftRpc::SetWindowSizeRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->SetWindowSize(request->hwnd(), request->width(), request->height()));
        }
    return ::grpc::Status();
}

::grpc::Status DmRpcImpl::SetWindowText(::grpc::ServerContext* context, const::dmsoftRpc::SetWindowTextRequest* request, ::dmsoftRpc::IsOkResponse* response)
{
    auto dm = getDmsoft(request->index());
    if (dm) {
        response->set_isok(dm->SetWindowText(request->hwnd(), string2Wstring(request->title())));
        }
    return ::grpc::Status();
}



