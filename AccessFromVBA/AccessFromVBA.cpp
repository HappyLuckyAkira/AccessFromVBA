// AccessFromVBA.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"
#include "AccessFromVBA.h"
#include <string>
#include <comutil.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 唯一のアプリケーション オブジェクトです。

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	// MFC を初期化して、エラーの場合は結果を印刷します。
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		// TODO: 必要に応じてエラー コードを変更してください。
		_tprintf(_T("致命的なエラー: MFC の初期化ができませんでした。\n"));
		nRetCode = 1;
	}
	else
	{
		// TODO: アプリケーションの動作を記述するコードをここに挿入してください。
	}

	return nRetCode;
}

ACCESSFROMVBA_API void WINAPI DoNothing()
{
    return;
}

//int型の値を受け取って、int型の戻り値を返す
ACCESSFROMVBA_API int WINAPI ReturnInt_Multi2(int i)
{
	return i * 2;
}

//int型の値を受け取って、内部で値を変更して返す
ACCESSFROMVBA_API void WINAPI ReturnIntByPointerParameter(int* pi)
{
	*pi *= 300;

	return;
}


std::wstring	convVstr2Wstr(const VARIANT v)
{
	std::wstring ws;

	if (v.vt == VT_BSTR)
	{
		ws = std::wstring(v.bstrVal, SysStringLen(v.bstrVal));
	}
	else if (v.vt == (VT_BSTR | VT_BYREF))
	{
		ws = std::wstring(*v.pbstrVal, SysStringLen(*v.pbstrVal));
	}

	return ws;
}

ACCESSFROMVBA_API void WINAPI SetString(VARIANT vString)
{
	if ((vString.vt == VT_BSTR) || (vString.vt == (VT_BSTR | VT_BYREF)))
	{
		MessageBox(NULL, convVstr2Wstr(vString).c_str(), L"DLL", MB_OK | MB_ICONINFORMATION);
	}
	else
	{
		MessageBox(NULL, L"No String Arg.", L"DLL", MB_OK | MB_ICONERROR);
	}

	return;
}

ACCESSFROMVBA_API void WINAPI GetStringByParam(VARIANT* pvString)
{
	if (!pvString)
	{
		return;
	}
	
	std::wstring ws(L"GetStringByParam返却データ");
	
	unsigned short vt = pvString->vt;

	VariantClear(pvString);

	if (vt == (VT_BSTR | VT_BYREF))
	{
		//VBAから、
		//GetStringByParam(String)
		//で呼ばれた場合
		pvString->vt = vt;

		BSTR bs = SysAllocString(ws.c_str());
		INT iResult = SysReAllocString(pvString->pbstrVal, bs);
		SysFreeString(bs);
	}
	else if ((vt == VT_BSTR) || (vt == VT_EMPTY))
	{
		//VBAから、
		//GetStringByParam(Variant)
		//で呼ばれた場合
		pvString->vt = VT_BSTR;
		pvString->bstrVal = SysAllocString(ws.c_str());

	}

	return;
}

ACCESSFROMVBA_API VARIANT WINAPI GetStringByRetVal()
{
	std::wstring ws(L"GetStringByRetVal");

	VARIANT vstr;

	VariantInit(&vstr);
	vstr.vt = VT_BSTR;
	vstr.bstrVal = SysAllocString(ws.c_str());

	return vstr;
}

ACCESSFROMVBA_API void WINAPI GetStringByParamS(BSTR* pbstr)
{
    if (!pbstr)
        return;

    //まず開放
    SysFreeString(*pbstr);

    //返す文字列（std::wstringは使用しない）
    std::string sReturn("GetStringByParamS 返却データ文字列");

    //BSTR生成
    *pbstr = SysAllocStringByteLen(sReturn.c_str(), sReturn.length());

    return;
}

ACCESSFROMVBA_API BSTR WINAPI GetStringByRetValS()
{
    //返す文字列（std::wstringは使用しない）
    std::string s("GetStringByRetValS 返却データ文字列");

    //BSTR生成
    BSTR bstr = SysAllocStringByteLen(s.c_str(), s.length());

    return bstr;
}

ACCESSFROMVBA_API void WINAPI SetStringS(const BSTR sString)
{
    if (!sString)
    {
        MessageBox(NULL, L"Argment is NULL.", L"DLL", MB_OK | MB_ICONERROR);

        return;
    }

    //VBAからの文字列は、char*で格納されている
    //BSTRの途中に'\0'がある事は想定していない。
    std::string s((char*)sString);

    MessageBoxA(NULL, s.c_str(), "DLL", MB_OK | MB_ICONINFORMATION);

    return;
}