//#include "stdafx.h"

#include "dmObj.h"

class MyDispatchDriver
{
public:
	IDispatch* p;
	MyDispatchDriver()
	{
		p = NULL;
	}
	MyDispatchDriver(IDispatch* lp)
	{
		if ((p = lp) != NULL)
			p->AddRef();//增加引用计数
	}
	~MyDispatchDriver() { if (p) p->Release(); }
	HRESULT GetIDOfName(LPCOLESTR lpsz, DISPID* pdispid)
	{
		HRESULT hr = -1;
		if (p == NULL) return hr;
		return p->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpsz, 1, LOCALE_USER_DEFAULT, pdispid);
	}
	HRESULT Invoke0(DISPID dispid, VARIANT* pvarRet = NULL)
	{
		HRESULT hr = -1;
		DISPPARAMS dispparams = { NULL, NULL, 0, 0 };
		if (p == NULL) return hr;
		return p->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pvarRet, NULL, NULL);
	}
	HRESULT InvokeN(DISPID dispid, VARIANT* pvarParams, int nParams, VARIANT* pvarRet = NULL)
	{
		HRESULT hr = -1;
		DISPPARAMS dispparams = { pvarParams, NULL, nParams, 0 };
		if (p == NULL) return hr;
		return p->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pvarRet, NULL, NULL);
	}
};

dmsoft::dmsoft()
{
	CLSID clsid;
	IUnknown* pUnknown = NULL;
	HRESULT hr;

	obj = NULL;
	clsid = { 0x26037A0E, 0x7CBD , 0x4FFF ,{ 0x9C,0x63 , 0x56,0xF2,0xD0,0x77,0x02,0x14 } };
	hr = ::CLSIDFromProgID(L"dm.dmsoft", &clsid);
	if (FAILED(hr))
	{
		return;
	}
	//CoInitialize(NULL);
	CoInitializeEx(NULL, 0);
	hr = ::CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IUnknown, (LPVOID*)&pUnknown);
	if (FAILED(hr))
	{
		return;
	}

	pUnknown->QueryInterface(IID_IDispatch, (void**)&obj);
	if (pUnknown) pUnknown->Release();
}

dmsoft::~dmsoft()
{
	if (obj) obj->Release();
	//CoUninitialize();
}

wstring dmsoft::Ver()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Ver", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetPath(const wstring path)
//long dmsoft::SetPath(std::wstring path)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(path.c_str());
	//pn[0] = COleVariant(path.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetPath", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::Ocr(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Ocr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

std::tuple<bool,int, int> dmsoft::FindStr(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(str.c_str());
	pn[3] = COleVariant(color.c_str());
	pn[2] = COleVariant(sim);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		return { vResult.lVal==1, t0.lVal, t1.lVal };
	}
	return { false,-1,-1 };
}

long dmsoft::GetResultCount(const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetResultCount", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::GetResultPos(const wstring str, long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[3] = COleVariant(str.c_str());
	pn[2] = COleVariant(index);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetResultPos", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return { vResult.lVal, t0.lVal, t1.lVal };
	}
	return { 0,-1,-1 };
}

long dmsoft::StrStr(const wstring s, const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(s.c_str());
	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"StrStr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SendCommand(const wstring cmd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(cmd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendCommand", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::UseDict(long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(index);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"UseDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetBasePath()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetBasePath", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetDictPwd(const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDictPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::OcrInFile(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(pic_name.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"OcrInFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::Capture(long x1, long y1, long x2, long y2, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(x1);
	pn[3] = COleVariant(y1);
	pn[2] = COleVariant(x2);
	pn[1] = COleVariant(y2);
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Capture", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyPress(long vk)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(vk);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyPress", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyDown(long vk)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(vk);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyDown", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyUp(long vk)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(vk);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyUp", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LeftClick()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LeftClick", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RightClick()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RightClick", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MiddleClick()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MiddleClick", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LeftDoubleClick()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LeftDoubleClick", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LeftDown()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LeftDown", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LeftUp()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LeftUp", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RightDown()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RightDown", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RightUp()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RightUp", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MoveTo(long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveTo", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MoveR(long rx, long ry)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(rx);
	pn[0] = COleVariant(ry);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveR", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetColor(long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetColor", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetColorBGR(long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetColorBGR", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::RGB2BGR(const wstring rgb_color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(rgb_color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RGB2BGR", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::BGR2RGB(const wstring bgr_color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(bgr_color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"BGR2RGB", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::UnBindWindow()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"UnBindWindow", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CmpColor(long x, long y, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CmpColor", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::ClientToScreen(long hwnd, long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[2] = COleVariant(hwnd);
	t0.vt = VT_I4;
	t0.lVal = x;
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	t1.vt = VT_I4;
	t1.lVal = y;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ClientToScreen", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		x = t0.lVal;
		y = t1.lVal;
		return { vResult.lVal,x,y };
	}
	return { 0,-1,-1 };
}

std::tuple<long, int, int> dmsoft::ScreenToClient(long hwnd, long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[2] = COleVariant(hwnd);
	t0.vt = VT_I4;
	t0.lVal = x;
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	t1.vt = VT_I4;
	t1.lVal = y;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ScreenToClient", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		x = t0.lVal;
		y = t1.lVal;
		return { vResult.lVal,x,y };
	}
	return { 0,-1,-1 };
}

long dmsoft::ShowScrMsg(long x1, long y1, long x2, long y2, const wstring msg, const wstring color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(msg.c_str());
	pn[0] = COleVariant(color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ShowScrMsg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetMinRowGap(long row_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(row_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetMinRowGap", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetMinColGap(long col_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(col_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetMinColGap", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::FindColor(long x1, long y1, long x2, long y2, const wstring color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindColor", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		return {vResult.lVal,t0.lVal,t1.lVal};
	}
	return { 0,-1,-1 };
}

wstring dmsoft::FindColorEx(long x1, long y1, long x2, long y2, const wstring color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindColorEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetWordLineHeight(long line_height)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(line_height);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWordLineHeight", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetWordGap(long word_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(word_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWordGap", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetRowGapNoDict(long row_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(row_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetRowGapNoDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetColGapNoDict(long col_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(col_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetColGapNoDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetWordLineHeightNoDict(long line_height)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(line_height);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWordLineHeightNoDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetWordGapNoDict(long word_gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(word_gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWordGapNoDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetWordResultCount(const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWordResultCount", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::GetWordResultPos(const wstring str, long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[3] = COleVariant(str.c_str());
	pn[2] = COleVariant(index);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWordResultPos", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return { vResult.lVal, t0.lVal,t1.lVal };
	}
	return { 0,-1,-1 };
}

wstring dmsoft::GetWordResultStr(const wstring str, long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(str.c_str());
	pn[0] = COleVariant(index);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWordResultStr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetWords(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWords", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetWordsNoDict(long x1, long y1, long x2, long y2, const wstring color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(x1);
	pn[3] = COleVariant(y1);
	pn[2] = COleVariant(x2);
	pn[1] = COleVariant(y2);
	pn[0] = COleVariant(color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWordsNoDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetShowErrorMsg(long show)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(show);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetShowErrorMsg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::GetClientSize(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[2] = COleVariant(hwnd);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetClientSize", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		auto width = t0.lVal;
		auto height = t1.lVal;
		return { vResult.lVal,width,height };
	}
	return { 0,-1,-1 };
}

long dmsoft::MoveWindow(long hwnd, long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetColorHSV(long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetColorHSV", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetAveRGB(long x1, long y1, long x2, long y2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x1);
	pn[2] = COleVariant(y1);
	pn[1] = COleVariant(x2);
	pn[0] = COleVariant(y2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetAveRGB", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetAveHSV(long x1, long y1, long x2, long y2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x1);
	pn[2] = COleVariant(y1);
	pn[1] = COleVariant(x2);
	pn[0] = COleVariant(y2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetAveHSV", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetForegroundWindow()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetForegroundWindow", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetForegroundFocus()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetForegroundFocus", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetMousePointWindow()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetMousePointWindow", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetPointWindow(long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetPointWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::EnumWindow(long parent, const wstring title, const wstring class_name, long filter)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(parent);
	pn[2] = COleVariant(title.c_str());
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(filter);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetWindowState(long hwnd, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowState", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetWindow(long hwnd, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetSpecialWindow(long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetSpecialWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetWindowText(long hwnd, const wstring text)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(text.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWindowText", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetWindowSize(long hwnd, long width, long height)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(width);
	pn[0] = COleVariant(height);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWindowSize", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int, int, int>  dmsoft::GetWindowRect(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;
	CComVariant t0, t1, t2, t3;

	VariantInit(&t0);
	VariantInit(&t1);
	VariantInit(&t2);
	VariantInit(&t3);

	pn[4] = COleVariant(hwnd);
	pn[3].vt = VT_BYREF | VT_VARIANT;
	pn[3].pvarVal = &t0;
	pn[2].vt = VT_BYREF | VT_VARIANT;
	pn[2].pvarVal = &t1;
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t2;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t3;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		int x1 = t0.lVal;
		int y1 = t1.lVal;
		int x2 = t2.lVal;
		int y2 = t3.lVal;
		return { vResult.lVal,x1,y1,x2,y2 };
	}
	return { 0,-1,-1,-1,-1 };
}

wstring dmsoft::GetWindowTitle(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowTitle", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetWindowClass(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowClass", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetWindowState(long hwnd, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWindowState", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CreateFoobarRect(long hwnd, long x, long y, long w, long h)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(w);
	pn[0] = COleVariant(h);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CreateFoobarRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CreateFoobarRoundRect(long hwnd, long x, long y, long w, long h, long rw, long rh)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(hwnd);
	pn[5] = COleVariant(x);
	pn[4] = COleVariant(y);
	pn[3] = COleVariant(w);
	pn[2] = COleVariant(h);
	pn[1] = COleVariant(rw);
	pn[0] = COleVariant(rh);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CreateFoobarRoundRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CreateFoobarEllipse(long hwnd, long x, long y, long w, long h)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(w);
	pn[0] = COleVariant(h);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CreateFoobarEllipse", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CreateFoobarCustom(long hwnd, long x, long y, const wstring pic, const wstring trans_color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(hwnd);
	pn[4] = COleVariant(x);
	pn[3] = COleVariant(y);
	pn[2] = COleVariant(pic.c_str());
	pn[1] = COleVariant(trans_color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CreateFoobarCustom", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarFillRect(long hwnd, long x1, long y1, long x2, long y2, const wstring color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(hwnd);
	pn[4] = COleVariant(x1);
	pn[3] = COleVariant(y1);
	pn[2] = COleVariant(x2);
	pn[1] = COleVariant(y2);
	pn[0] = COleVariant(color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarFillRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarDrawText(long hwnd, long x, long y, long w, long h, const wstring text, const wstring color, long align)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(hwnd);
	pn[6] = COleVariant(x);
	pn[5] = COleVariant(y);
	pn[4] = COleVariant(w);
	pn[3] = COleVariant(h);
	pn[2] = COleVariant(text.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(align);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarDrawText", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarDrawPic(long hwnd, long x, long y, const wstring pic, const wstring trans_color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(pic.c_str());
	pn[0] = COleVariant(trans_color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarDrawPic", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarUpdate(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarUpdate", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarLock(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarLock", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarUnlock(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarUnlock", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarSetFont(long hwnd, const wstring font_name, long size, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(font_name.c_str());
	pn[1] = COleVariant(size);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarSetFont", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarTextRect(long hwnd, long x, long y, long w, long h)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(w);
	pn[0] = COleVariant(h);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarTextRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarPrintText(long hwnd, const wstring text, const wstring color)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(text.c_str());
	pn[0] = COleVariant(color.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarPrintText", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarClearText(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarClearText", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarTextLineGap(long hwnd, long gap)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(gap);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarTextLineGap", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Play(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Play", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqCapture(long x1, long y1, long x2, long y2, long quality, long delay, long time)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(quality);
	pn[1] = COleVariant(delay);
	pn[0] = COleVariant(time);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqCapture", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqRelease(long handle)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(handle);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqRelease", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FaqSend(const wstring server, long handle, long request_type, long time_out)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(server.c_str());
	pn[2] = COleVariant(handle);
	pn[1] = COleVariant(request_type);
	pn[0] = COleVariant(time_out);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqSend", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::Beep(long fre, long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(fre);
	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Beep", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarClose(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarClose", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MoveDD(long dx, long dy)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(dx);
	pn[0] = COleVariant(dy);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveDD", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqGetSize(long handle)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(handle);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqGetSize", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LoadPic(const wstring pic_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pic_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LoadPic", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FreePic(const wstring pic_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pic_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FreePic", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetScreenData(long x1, long y1, long x2, long y2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x1);
	pn[2] = COleVariant(y1);
	pn[1] = COleVariant(x2);
	pn[0] = COleVariant(y2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetScreenData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FreeScreenData(long handle)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(handle);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FreeScreenData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WheelUp()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WheelUp", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WheelDown()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WheelDown", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetMouseDelay(const wstring tpe, long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(tpe.c_str());
	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetMouseDelay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetKeypadDelay(const wstring tpe, long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(tpe.c_str());
	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetKeypadDelay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetEnv(long index, const wstring name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetEnv", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetEnv(long index, const wstring name, const wstring value)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(index);
	pn[1] = COleVariant(name.c_str());
	pn[0] = COleVariant(value.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetEnv", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SendString(long hwnd, const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendString", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DelEnv(long index, const wstring name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DelEnv", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetPath()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetPath", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetDict(long index, const wstring dict_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(dict_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<bool, int, int> dmsoft::FindPic(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring delta_color, double sim, long dir)
{
	long x; long y;
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(pic_name.c_str());
	pn[4] = COleVariant(delta_color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPic", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		x = t0.lVal;
		y = t1.lVal;
		return { vResult.lVal==1,x,y };
	}
	return { false,-1,-1 };
}

wstring dmsoft::FindPicEx(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring delta_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(pic_name.c_str());
	pn[2] = COleVariant(delta_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetClientSize(long hwnd, long width, long height)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(width);
	pn[0] = COleVariant(height);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetClientSize", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::ReadInt(long hwnd, const wstring addr, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr.c_str());
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadInt", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

float dmsoft::ReadFloat(long hwnd, const wstring addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(addr.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadFloat", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.fltVal;
	}
	return 0.0f;
}

double dmsoft::ReadDouble(long hwnd, const wstring addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(addr.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadDouble", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.dblVal;
	}
	return 0.0;
}

wstring dmsoft::FindInt(long hwnd, const wstring addr_range, long int_value_min, long int_value_max, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(addr_range.c_str());
	pn[2] = COleVariant(int_value_min);
	pn[1] = COleVariant(int_value_max);
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindInt", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindFloat(long hwnd, const wstring addr_range, float float_value_min, float float_value_max)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr_range.c_str());
	pn[1] = COleVariant(float_value_min);
	pn[0] = COleVariant(float_value_max);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindFloat", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindDouble(long hwnd, const wstring addr_range, double double_value_min, double double_value_max)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr_range.c_str());
	pn[1] = COleVariant(double_value_min);
	pn[0] = COleVariant(double_value_max);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindDouble", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindString(long hwnd, const wstring addr_range, const wstring string_value, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr_range.c_str());
	pn[1] = COleVariant(string_value.c_str());
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindString", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetModuleBaseAddr(long hwnd, const wstring module_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(module_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetModuleBaseAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::MoveToEx(long x, long y, long w, long h)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x);
	pn[2] = COleVariant(y);
	pn[1] = COleVariant(w);
	pn[0] = COleVariant(h);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveToEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::MatchPicName(const wstring pic_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pic_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MatchPicName", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::AddDict(long index, const wstring dict_info)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(dict_info.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AddDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnterCri()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnterCri", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LeaveCri()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LeaveCri", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteInt(long hwnd, const wstring addr, long tpe, long v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr.c_str());
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteInt", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteFloat(long hwnd, const wstring addr, float v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr.c_str());
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteFloat", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteDouble(long hwnd, const wstring addr, double v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr.c_str());
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteDouble", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteString(long hwnd, const wstring addr, long tpe, const wstring v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr.c_str());
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(v.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteString", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::AsmAdd(const wstring asm_ins)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(asm_ins.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AsmAdd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::AsmClear()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AsmClear", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::AsmCall(long hwnd, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AsmCall", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::FindMultiColor(long x1, long y1, long x2, long y2, const wstring first_color, const wstring offset_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(first_color.c_str());
	pn[4] = COleVariant(offset_color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindMultiColor", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return { vResult.lVal,x,y };
	}
	return { 0,-1,-1 };
}

wstring dmsoft::FindMultiColorEx(long x1, long y1, long x2, long y2, const wstring first_color, const wstring offset_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(first_color.c_str());
	pn[2] = COleVariant(offset_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindMultiColorEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::AsmCode(long base_addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(base_addr);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AsmCode", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::Assemble(const wstring asm_code, long base_addr, long is_upper)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(asm_code.c_str());
	pn[1] = COleVariant(base_addr);
	pn[0] = COleVariant(is_upper);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Assemble", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetWindowTransparent(long hwnd, long v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetWindowTransparent", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadData(long hwnd, const wstring addr, long length)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr.c_str());
	pn[0] = COleVariant(length);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WriteData(long hwnd, const wstring addr, const wstring data)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr.c_str());
	pn[0] = COleVariant(data.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindData(long hwnd, const wstring addr_range, const wstring data)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr_range.c_str());
	pn[0] = COleVariant(data.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetPicPwd(const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetPicPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Log(const wstring info)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(info.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Log", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindStrE(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindColorE(long x1, long y1, long x2, long y2, const wstring color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindColorE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindPicE(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring delta_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(pic_name.c_str());
	pn[2] = COleVariant(delta_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindMultiColorE(long x1, long y1, long x2, long y2, const wstring first_color, const wstring offset_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(first_color.c_str());
	pn[2] = COleVariant(offset_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindMultiColorE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetExactOcr(long exact_ocr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(exact_ocr);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetExactOcr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadString(long hwnd, const wstring addr, long tpe, long length)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr.c_str());
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(length);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadString", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::FoobarTextPrintDir(long hwnd, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarTextPrintDir", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::OcrEx(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"OcrEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetDisplayInput(const wstring mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(mode.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDisplayInput", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetTime()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetTime", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetScreenWidth()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetScreenWidth", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetScreenHeight()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetScreenHeight", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::BindWindowEx(long hwnd, const wstring display, const wstring mouse, const wstring keypad, const wstring public_desc, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(hwnd);
	pn[4] = COleVariant(display.c_str());
	pn[3] = COleVariant(mouse.c_str());
	pn[2] = COleVariant(keypad.c_str());
	pn[1] = COleVariant(public_desc.c_str());
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"BindWindowEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetDiskSerial()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDiskSerial", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::Md5(const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Md5", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetMac()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetMac", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::ActiveInputMethod(long hwnd, const wstring id)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(id.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ActiveInputMethod", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CheckInputMethod(long hwnd, const wstring id)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(id.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CheckInputMethod", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindInputMethod(const wstring id)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(id.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindInputMethod", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::GetCursorPos()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetCursorPos", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return { vResult.lVal,x,y };
	}
	return { 0,-1,-1 };
}

long dmsoft::BindWindow(long hwnd, const wstring display, const wstring mouse, const wstring keypad, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(display.c_str());
	pn[2] = COleVariant(mouse.c_str());
	pn[1] = COleVariant(keypad.c_str());
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"BindWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindWindow(const wstring class_name, const wstring title_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(title_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetScreenDepth()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetScreenDepth", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetScreen(long width, long height, long depth)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(width);
	pn[1] = COleVariant(height);
	pn[0] = COleVariant(depth);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetScreen", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::ExitOs(long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ExitOs", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetDir(long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDir", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetOsType()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetOsType", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindWindowEx(long parent, const wstring class_name, const wstring title_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(parent);
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(title_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindWindowEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetExportDict(long index, const wstring dict_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(dict_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetExportDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetCursorShape()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetCursorShape", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::DownCpu(long rate)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(rate);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DownCpu", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetCursorSpot()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetCursorSpot", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SendString2(long hwnd, const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendString2", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqPost(const wstring server, long handle, long request_type, long time_out)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(server.c_str());
	pn[2] = COleVariant(handle);
	pn[1] = COleVariant(request_type);
	pn[0] = COleVariant(time_out);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqPost", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FaqFetch()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqFetch", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FetchWord(long x1, long y1, long x2, long y2, const wstring color, const wstring word)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(word.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FetchWord", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::CaptureJpg(long x1, long y1, long x2, long y2, const wstring file_name, long quality)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(quality);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CaptureJpg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::FindStrWithFont(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim, const wstring font_name, long font_size, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[12];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[11] = COleVariant(x1);
	pn[10] = COleVariant(y1);
	pn[9] = COleVariant(x2);
	pn[8] = COleVariant(y2);
	pn[7] = COleVariant(str.c_str());
	pn[6] = COleVariant(color.c_str());
	pn[5] = COleVariant(sim);
	pn[4] = COleVariant(font_name.c_str());
	pn[3] = COleVariant(font_size);
	pn[2] = COleVariant(flag);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrWithFont", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 12, &vResult);
	if (SUCCEEDED(hr))
	{
		 auto x = t0.lVal;
		auto y = t1.lVal;
		return { vResult.lVal,x,y };
	}
	return { 0,-1,-1 };
}

wstring dmsoft::FindStrWithFontE(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim, const wstring font_name, long font_size, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(str.c_str());
	pn[4] = COleVariant(color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(font_name.c_str());
	pn[1] = COleVariant(font_size);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrWithFontE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrWithFontEx(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim, const wstring font_name, long font_size, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(str.c_str());
	pn[4] = COleVariant(color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(font_name.c_str());
	pn[1] = COleVariant(font_size);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrWithFontEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetDictInfo(const wstring str, const wstring font_name, long font_size, long flag)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(str.c_str());
	pn[2] = COleVariant(font_name.c_str());
	pn[1] = COleVariant(font_size);
	pn[0] = COleVariant(flag);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDictInfo", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SaveDict(long index, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SaveDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetWindowProcessId(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowProcessId", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetWindowProcessPath(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetWindowProcessPath", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::LockInput(long locks)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(locks);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LockInput", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetPicSize(const wstring pic_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pic_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetPicSize", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetID()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetID", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CapturePng(long x1, long y1, long x2, long y2, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(x1);
	pn[3] = COleVariant(y1);
	pn[2] = COleVariant(x2);
	pn[1] = COleVariant(y2);
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CapturePng", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CaptureGif(long x1, long y1, long x2, long y2, const wstring file_name, long delay, long time)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(file_name.c_str());
	pn[1] = COleVariant(delay);
	pn[0] = COleVariant(time);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CaptureGif", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::ImageToBmp(const wstring pic_name, const wstring bmp_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(pic_name.c_str());
	pn[0] = COleVariant(bmp_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ImageToBmp", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

std::tuple<long, int, int> dmsoft::FindStrFast(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(str.c_str());
	pn[3] = COleVariant(color.c_str());
	pn[2] = COleVariant(sim);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrFast", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return { vResult.lVal,-1,-1 };
	}
	return {0, -1, -1};
}

wstring dmsoft::FindStrFastEx(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrFastEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrFastE(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrFastE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::EnableDisplayDebug(long enable_debug)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(enable_debug);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableDisplayDebug", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CapturePre(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CapturePre", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RegEx(const wstring code, const wstring Ver, const wstring ip)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(code.c_str());
	pn[1] = COleVariant(Ver.c_str());
	pn[0] = COleVariant(ip.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RegEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetMachineCode()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetMachineCode", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetClipboard(const wstring data)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(data.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetClipboard", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetClipboard()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetClipboard", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetNowDict()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetNowDict", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Is64Bit()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Is64Bit", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetColorNum(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetColorNum", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::EnumWindowByProcess(const wstring process_name, const wstring title, const wstring class_name, long filter)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(process_name.c_str());
	pn[2] = COleVariant(title.c_str());
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(filter);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumWindowByProcess", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetDictCount(long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(index);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDictCount", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetLastError()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetLastError", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetNetTime()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetNetTime", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::EnableGetColorByCapture(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableGetColorByCapture", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CheckUAC()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CheckUAC", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetUAC(long uac)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(uac);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetUAC", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DisableFontSmooth()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DisableFontSmooth", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CheckFontSmooth()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CheckFontSmooth", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetDisplayAcceler(long level)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(level);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDisplayAcceler", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindWindowByProcess(const wstring process_name, const wstring class_name, const wstring title_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(process_name.c_str());
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(title_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindWindowByProcess", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindWindowByProcessId(long process_id, const wstring class_name, const wstring title_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(process_id);
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(title_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindWindowByProcessId", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadIni(const wstring section, const wstring key, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(section.c_str());
	pn[1] = COleVariant(key.c_str());
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadIni", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WriteIni(const wstring section, const wstring key, const wstring v, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(section.c_str());
	pn[2] = COleVariant(key.c_str());
	pn[1] = COleVariant(v.c_str());
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteIni", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RunApp(const wstring path, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(path.c_str());
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RunApp", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Delay(long mis)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(mis);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"delay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindWindowSuper(const wstring spec1, long flag1, long type1, const wstring spec2, long flag2, long type2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(spec1.c_str());
	pn[4] = COleVariant(flag1);
	pn[3] = COleVariant(type1);
	pn[2] = COleVariant(spec2.c_str());
	pn[1] = COleVariant(flag2);
	pn[0] = COleVariant(type2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindWindowSuper", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ExcludePos(const wstring all_pos, long tpe, long x1, long y1, long x2, long y2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(all_pos.c_str());
	pn[4] = COleVariant(tpe);
	pn[3] = COleVariant(x1);
	pn[2] = COleVariant(y1);
	pn[1] = COleVariant(x2);
	pn[0] = COleVariant(y2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ExcludePos", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindNearestPos(const wstring all_pos, long tpe, long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(all_pos.c_str());
	pn[2] = COleVariant(tpe);
	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindNearestPos", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::SortPosDistance(const wstring all_pos, long tpe, long x, long y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(all_pos.c_str());
	pn[2] = COleVariant(tpe);
	pn[1] = COleVariant(x);
	pn[0] = COleVariant(y);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SortPosDistance", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::FindPicMem(long x1, long y1, long x2, long y2, const wstring pic_info, const wstring delta_color, double sim, long dir, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(pic_info.c_str());
	pn[4] = COleVariant(delta_color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicMem", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindPicMemEx(long x1, long y1, long x2, long y2, const wstring pic_info, const wstring delta_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(pic_info.c_str());
	pn[2] = COleVariant(delta_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicMemEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindPicMemE(long x1, long y1, long x2, long y2, const wstring pic_info, const wstring delta_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(pic_info.c_str());
	pn[2] = COleVariant(delta_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicMemE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::AppendPicAddr(const wstring pic_info, long addr, long size)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(pic_info.c_str());
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(size);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"AppendPicAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WriteFile(const wstring file_name, const wstring content)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(content.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Stop(long id)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(id);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Stop", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetDictMem(long index, long addr, long size)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(index);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(size);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDictMem", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetNetTimeSafe()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetNetTimeSafe", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::ForceUnBindWindow(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ForceUnBindWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadIniPwd(const wstring section, const wstring key, const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(section.c_str());
	pn[2] = COleVariant(key.c_str());
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadIniPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WriteIniPwd(const wstring section, const wstring key, const wstring v, const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(section.c_str());
	pn[3] = COleVariant(key.c_str());
	pn[2] = COleVariant(v.c_str());
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteIniPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DecodeFile(const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DecodeFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyDownChar(const wstring key_str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(key_str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyDownChar", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyUpChar(const wstring key_str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(key_str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyUpChar", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyPressChar(const wstring key_str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(key_str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyPressChar", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::KeyPressStr(const wstring key_str, long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(key_str.c_str());
	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"KeyPressStr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableKeypadPatch(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableKeypadPatch", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableKeypadSync(long en, long time_out)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(en);
	pn[0] = COleVariant(time_out);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableKeypadSync", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableMouseSync(long en, long time_out)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(en);
	pn[0] = COleVariant(time_out);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableMouseSync", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DmGuard(long en, const wstring tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(en);
	pn[0] = COleVariant(tpe.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DmGuard", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqCaptureFromFile(long x1, long y1, long x2, long y2, const wstring file_name, long quality)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(quality);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqCaptureFromFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindIntEx(long hwnd, const wstring addr_range, long int_value_min, long int_value_max, long tpe, long steps, long multi_thread, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(hwnd);
	pn[6] = COleVariant(addr_range.c_str());
	pn[5] = COleVariant(int_value_min);
	pn[4] = COleVariant(int_value_max);
	pn[3] = COleVariant(tpe);
	pn[2] = COleVariant(steps);
	pn[1] = COleVariant(multi_thread);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindIntEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindFloatEx(long hwnd, const wstring addr_range, float float_value_min, float float_value_max, long steps, long multi_thread, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(hwnd);
	pn[5] = COleVariant(addr_range.c_str());
	pn[4] = COleVariant(float_value_min);
	pn[3] = COleVariant(float_value_max);
	pn[2] = COleVariant(steps);
	pn[1] = COleVariant(multi_thread);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindFloatEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindDoubleEx(long hwnd, const wstring addr_range, double double_value_min, double double_value_max, long steps, long multi_thread, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(hwnd);
	pn[5] = COleVariant(addr_range.c_str());
	pn[4] = COleVariant(double_value_min);
	pn[3] = COleVariant(double_value_max);
	pn[2] = COleVariant(steps);
	pn[1] = COleVariant(multi_thread);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindDoubleEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStringEx(long hwnd, const wstring addr_range, const wstring string_value, long tpe, long steps, long multi_thread, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(hwnd);
	pn[5] = COleVariant(addr_range.c_str());
	pn[4] = COleVariant(string_value.c_str());
	pn[3] = COleVariant(tpe);
	pn[2] = COleVariant(steps);
	pn[1] = COleVariant(multi_thread);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStringEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindDataEx(long hwnd, const wstring addr_range, const wstring data, long steps, long multi_thread, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(hwnd);
	pn[4] = COleVariant(addr_range.c_str());
	pn[3] = COleVariant(data.c_str());
	pn[2] = COleVariant(steps);
	pn[1] = COleVariant(multi_thread);
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindDataEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::EnableRealMouse(long en, long mousedelay, long mousestep)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(en);
	pn[1] = COleVariant(mousedelay);
	pn[0] = COleVariant(mousestep);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableRealMouse", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableRealKeypad(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableRealKeypad", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SendStringIme(const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendStringIme", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarDrawLine(long hwnd, long x1, long y1, long x2, long y2, const wstring color, long style, long width)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(hwnd);
	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(color.c_str());
	pn[1] = COleVariant(style);
	pn[0] = COleVariant(width);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarDrawLine", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindStrEx(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::IsBind(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"IsBind", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetDisplayDelay(long t)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(t);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetDisplayDelay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetDmCount()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDmCount", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DisableScreenSave()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DisableScreenSave", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DisablePowerSave()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DisablePowerSave", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetMemoryHwndAsProcessId(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetMemoryHwndAsProcessId", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindShape(long x1, long y1, long x2, long y2, const wstring offset_color, double sim, long dir, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(offset_color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindShape", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindShapeE(long x1, long y1, long x2, long y2, const wstring offset_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(offset_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindShapeE", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindShapeEx(long x1, long y1, long x2, long y2, const wstring offset_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(offset_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindShapeEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrS(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(str.c_str());
	pn[3] = COleVariant(color.c_str());
	pn[2] = COleVariant(sim);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrExS(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrExS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrFastS(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(str.c_str());
	pn[3] = COleVariant(color.c_str());
	pn[2] = COleVariant(sim);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrFastS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindStrFastExS(long x1, long y1, long x2, long y2, const wstring str, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(x1);
	pn[5] = COleVariant(y1);
	pn[4] = COleVariant(x2);
	pn[3] = COleVariant(y2);
	pn[2] = COleVariant(str.c_str());
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindStrFastExS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindPicS(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring delta_color, double sim, long dir, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[10];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[9] = COleVariant(x1);
	pn[8] = COleVariant(y1);
	pn[7] = COleVariant(x2);
	pn[6] = COleVariant(y2);
	pn[5] = COleVariant(pic_name.c_str());
	pn[4] = COleVariant(delta_color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(dir);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 10, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FindPicExS(long x1, long y1, long x2, long y2, const wstring pic_name, const wstring delta_color, double sim, long dir)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[8];
	CComVariant vResult;

	pn[7] = COleVariant(x1);
	pn[6] = COleVariant(y1);
	pn[5] = COleVariant(x2);
	pn[4] = COleVariant(y2);
	pn[3] = COleVariant(pic_name.c_str());
	pn[2] = COleVariant(delta_color.c_str());
	pn[1] = COleVariant(sim);
	pn[0] = COleVariant(dir);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindPicExS", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 8, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::ClearDict(long index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(index);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ClearDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetMachineCodeNoMac()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetMachineCodeNoMac", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

std::tuple<long, int, int, int, int> dmsoft::GetClientRect(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;
	CComVariant t0, t1, t2, t3;

	VariantInit(&t0);
	VariantInit(&t1);
	VariantInit(&t2);
	VariantInit(&t3);

	pn[4] = COleVariant(hwnd);
	pn[3].vt = VT_BYREF | VT_VARIANT;
	pn[3].pvarVal = &t0;
	pn[2].vt = VT_BYREF | VT_VARIANT;
	pn[2].pvarVal = &t1;
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t2;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t3;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetClientRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x1 = t0.lVal;
		auto y1 = t1.lVal;
		auto x2 = t2.lVal;
		auto y2 = t3.lVal;
		return { vResult.lVal,x1,y1,x2,y2 };
	}
	return { 0,-1,-1,-1,-1 };
}

long dmsoft::EnableFakeActive(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableFakeActive", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetScreenDataBmp(long x1, long y1, long x2, long y2, long* data, long* size)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetScreenDataBmp", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		*data = t0.lVal;
		*size = t1.lVal;
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EncodeFile(const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EncodeFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetCursorShapeEx(long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetCursorShapeEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::FaqCancel()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqCancel", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::IntToData(long int_value, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(int_value);
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"IntToData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::FloatToData(float float_value)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(float_value);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FloatToData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::DoubleToData(double double_value)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(double_value);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DoubleToData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::StringToData(const wstring string_value, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(string_value.c_str());
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"StringToData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetMemoryFindResultToFile(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetMemoryFindResultToFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableBind(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableBind", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetSimMode(long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetSimMode", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LockMouseRect(long x1, long y1, long x2, long y2)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(x1);
	pn[2] = COleVariant(y1);
	pn[1] = COleVariant(x2);
	pn[0] = COleVariant(y2);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LockMouseRect", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SendPaste(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendPaste", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::IsDisplayDead(long x1, long y1, long x2, long y2, long t)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(x1);
	pn[3] = COleVariant(y1);
	pn[2] = COleVariant(x2);
	pn[1] = COleVariant(y2);
	pn[0] = COleVariant(t);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"IsDisplayDead", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetKeyState(long vk)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(vk);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetKeyState", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CopyFile(const wstring src_file, const wstring dst_file, long over)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(src_file.c_str());
	pn[1] = COleVariant(dst_file.c_str());
	pn[0] = COleVariant(over);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CopyFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::IsFileExist(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"IsFileExist", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DeleteFile(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DeleteFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MoveFile(const wstring src_file, const wstring dst_file)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(src_file.c_str());
	pn[0] = COleVariant(dst_file.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MoveFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::CreateFolder(const wstring folder_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(folder_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"CreateFolder", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DeleteFolder(const wstring folder_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(folder_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DeleteFolder", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::GetFileLength(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetFileLength", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadFile(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WaitKey(long key_code, long time_out)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(key_code);
	pn[0] = COleVariant(time_out);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WaitKey", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DeleteIni(const wstring section, const wstring key, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(section.c_str());
	pn[1] = COleVariant(key.c_str());
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DeleteIni", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::DeleteIniPwd(const wstring section, const wstring key, const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(section.c_str());
	pn[2] = COleVariant(key.c_str());
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DeleteIniPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableSpeedDx(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableSpeedDx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableIme(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableIme", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::Reg(const wstring code, const wstring Ver)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(code.c_str());
	pn[0] = COleVariant(Ver.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Reg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::SelectFile()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SelectFile", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::SelectDirectory()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SelectDirectory", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::LockDisplay(long locks)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(locks);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LockDisplay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarSetSave(long hwnd, const wstring file_name, long en, const wstring header)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(file_name.c_str());
	pn[1] = COleVariant(en);
	pn[0] = COleVariant(header.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarSetSave", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::EnumWindowSuper(const wstring spec1, long flag1, long type1, const wstring spec2, long flag2, long type2, long sort)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[7];
	CComVariant vResult;

	pn[6] = COleVariant(spec1.c_str());
	pn[5] = COleVariant(flag1);
	pn[4] = COleVariant(type1);
	pn[3] = COleVariant(spec2.c_str());
	pn[2] = COleVariant(flag2);
	pn[1] = COleVariant(type2);
	pn[0] = COleVariant(sort);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumWindowSuper", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 7, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::DownloadFile(const wstring url, const wstring save_file, long timeout)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(url.c_str());
	pn[1] = COleVariant(save_file.c_str());
	pn[0] = COleVariant(timeout);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"DownloadFile", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableKeypadMsg(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableKeypadMsg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnableMouseMsg(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableMouseMsg", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RegNoMac(const wstring code, const wstring Ver)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(code.c_str());
	pn[0] = COleVariant(Ver.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RegNoMac", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::RegExNoMac(const wstring code, const wstring Ver, const wstring ip)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(code.c_str());
	pn[1] = COleVariant(Ver.c_str());
	pn[0] = COleVariant(ip.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"RegExNoMac", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SetEnumWindowDelay(long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetEnumWindowDelay", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindMulColor(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindMulColor", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetDict(long index, long font_index)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(index);
	pn[0] = COleVariant(font_index);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDict", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::GetBindWindow()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetBindWindow", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarStartGif(long hwnd, long x, long y, const wstring pic_name, long repeat_limit, long delay)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(hwnd);
	pn[4] = COleVariant(x);
	pn[3] = COleVariant(y);
	pn[2] = COleVariant(pic_name.c_str());
	pn[1] = COleVariant(repeat_limit);
	pn[0] = COleVariant(delay);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarStartGif", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarStopGif(long hwnd, long x, long y, const wstring pic_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(x);
	pn[1] = COleVariant(y);
	pn[0] = COleVariant(pic_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarStopGif", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FreeProcessMemory(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FreeProcessMemory", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadFileData(const wstring file_name, long start_pos, long end_pos)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(file_name.c_str());
	pn[1] = COleVariant(start_pos);
	pn[0] = COleVariant(end_pos);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadFileData", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::VirtualAllocEx(long hwnd, long addr, long size, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr);
	pn[1] = COleVariant(size);
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"VirtualAllocEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::VirtualFreeEx(long hwnd, long addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(addr);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"VirtualFreeEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetCommandLine(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetCommandLine", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::TerminateProcess(long pid)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pid);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"TerminateProcess", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetNetTimeByIp(const wstring ip)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(ip.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetNetTimeByIp", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::EnumProcess(const wstring name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumProcess", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetProcessInfo(long pid)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pid);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetProcessInfo", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::ReadIntAddr(long hwnd, long addr, long tpe)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(tpe);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadIntAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::ReadDataAddr(long hwnd, long addr, long length)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(length);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadDataAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

double dmsoft::ReadDoubleAddr(long hwnd, long addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(addr);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadDoubleAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.dblVal;
	}
	return 0.0;
}

float dmsoft::ReadFloatAddr(long hwnd, long addr)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(hwnd);
	pn[0] = COleVariant(addr);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadFloatAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.fltVal;
	}
	return 0.0f;
}

wstring dmsoft::ReadStringAddr(long hwnd, long addr, long tpe, long length)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr);
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(length);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"ReadStringAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::WriteDataAddr(long hwnd, long addr, const wstring data)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(data.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteDataAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteDoubleAddr(long hwnd, long addr, double v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteDoubleAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteFloatAddr(long hwnd, long addr, float v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(addr);
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteFloatAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteIntAddr(long hwnd, long addr, long tpe, long v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr);
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(v);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteIntAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::WriteStringAddr(long hwnd, long addr, long tpe, const wstring v)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(addr);
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(v.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"WriteStringAddr", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}
//long dmsoft::Delays(double min_s, double max_s){
//	Delays((int) min_s, (int)max_s);
//};
long dmsoft::Delays(long min_s, long max_s)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(min_s);
	pn[0] = COleVariant(max_s);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"Delays", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FindColorBlock(long x1, long y1, long x2, long y2, const wstring color, double sim, long count, long width, long height, long* x, long* y)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[11];
	CComVariant vResult;
	CComVariant t0, t1;

	VariantInit(&t0);
	VariantInit(&t1);

	pn[10] = COleVariant(x1);
	pn[9] = COleVariant(y1);
	pn[8] = COleVariant(x2);
	pn[7] = COleVariant(y2);
	pn[6] = COleVariant(color.c_str());
	pn[5] = COleVariant(sim);
	pn[4] = COleVariant(count);
	pn[3] = COleVariant(width);
	pn[2] = COleVariant(height);
	pn[1].vt = VT_BYREF | VT_VARIANT;
	pn[1].pvarVal = &t0;
	pn[0].vt = VT_BYREF | VT_VARIANT;
	pn[0].pvarVal = &t1;

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindColorBlock", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 11, &vResult);
	if (SUCCEEDED(hr))
	{
		auto x = t0.lVal;
		auto y = t1.lVal;
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::FindColorBlockEx(long x1, long y1, long x2, long y2, const wstring color, double sim, long count, long width, long height)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[9];
	CComVariant vResult;

	pn[8] = COleVariant(x1);
	pn[7] = COleVariant(y1);
	pn[6] = COleVariant(x2);
	pn[5] = COleVariant(y2);
	pn[4] = COleVariant(color.c_str());
	pn[3] = COleVariant(sim);
	pn[2] = COleVariant(count);
	pn[1] = COleVariant(width);
	pn[0] = COleVariant(height);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FindColorBlockEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 9, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::OpenProcess(long pid)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(pid);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"OpenProcess", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::EnumIniSection(const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumIniSection", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::EnumIniSectionPwd(const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumIniSectionPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::EnumIniKey(const wstring section, const wstring file_name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(section.c_str());
	pn[0] = COleVariant(file_name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumIniKey", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::EnumIniKeyPwd(const wstring section, const wstring file_name, const wstring pwd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(section.c_str());
	pn[1] = COleVariant(file_name.c_str());
	pn[0] = COleVariant(pwd.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumIniKeyPwd", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SwitchBindWindow(long hwnd)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(hwnd);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SwitchBindWindow", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::InitCri()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"InitCri", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::SendStringIme2(long hwnd, const wstring str, long mode)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(hwnd);
	pn[1] = COleVariant(str.c_str());
	pn[0] = COleVariant(mode);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SendStringIme2", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::EnumWindowByProcessId(long pid, const wstring title, const wstring class_name, long filter)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(pid);
	pn[2] = COleVariant(title.c_str());
	pn[1] = COleVariant(class_name.c_str());
	pn[0] = COleVariant(filter);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnumWindowByProcessId", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

wstring dmsoft::GetDisplayInfo()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetDisplayInfo", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::EnableFontSmooth()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnableFontSmooth", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::OcrExOne(long x1, long y1, long x2, long y2, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[6];
	CComVariant vResult;

	pn[5] = COleVariant(x1);
	pn[4] = COleVariant(y1);
	pn[3] = COleVariant(x2);
	pn[2] = COleVariant(y2);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"OcrExOne", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 6, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::SetAero(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"SetAero", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FoobarSetTrans(long hwnd, long trans, const wstring color, double sim)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[4];
	CComVariant vResult;

	pn[3] = COleVariant(hwnd);
	pn[2] = COleVariant(trans);
	pn[1] = COleVariant(color.c_str());
	pn[0] = COleVariant(sim);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FoobarSetTrans", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 4, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::EnablePicCache(long en)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(en);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"EnablePicCache", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

wstring dmsoft::GetInfo(const wstring cmd, const wstring param)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[2];
	CComVariant vResult;

	pn[1] = COleVariant(cmd.c_str());
	pn[0] = COleVariant(param.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"GetInfo", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 2, &vResult);
	if (SUCCEEDED(hr))
	{
		return  wstring(vResult.bstrVal);
	}
	return  _T("");
}

long dmsoft::FaqIsPosted()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqIsPosted", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::LoadPicByte(long addr, long size, const wstring name)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[3];
	CComVariant vResult;

	pn[2] = COleVariant(addr);
	pn[1] = COleVariant(size);
	pn[0] = COleVariant(name.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"LoadPicByte", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 3, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MiddleDown()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MiddleDown", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::MiddleUp()
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	CComVariant vResult;


	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"MiddleUp", &dispatch_id);
	}

	hr = spDisp.Invoke0(dispatch_id, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::FaqCaptureString(const wstring str)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[1];
	CComVariant vResult;

	pn[0] = COleVariant(str.c_str());

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"FaqCaptureString", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 1, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

long dmsoft::VirtualProtectEx(long hwnd, long addr, long size, long tpe, long old_protect)
{
	static DISPID dispatch_id = -1;
	MyDispatchDriver spDisp(obj);
	HRESULT hr;
	COleVariant pn[5];
	CComVariant vResult;

	pn[4] = COleVariant(hwnd);
	pn[3] = COleVariant(addr);
	pn[2] = COleVariant(size);
	pn[1] = COleVariant(tpe);
	pn[0] = COleVariant(old_protect);

	if (dispatch_id == -1)
	{
		spDisp.GetIDOfName(L"VirtualProtectEx", &dispatch_id);
	}

	hr = spDisp.InvokeN(dispatch_id, pn, 5, &vResult);
	if (SUCCEEDED(hr))
	{
		return vResult.lVal;
	}
	return 0;
}

