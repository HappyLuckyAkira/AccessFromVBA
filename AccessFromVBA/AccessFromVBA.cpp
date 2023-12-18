// AccessFromVBA.cpp : DLL �A�v���P�[�V�����p�ɃG�N�X�|�[�g�����֐����`���܂��B
//

#include "stdafx.h"
#include "AccessFromVBA.h"
#include <string>
#include <comutil.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// �B��̃A�v���P�[�V���� �I�u�W�F�N�g�ł��B

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	// MFC �����������āA�G���[�̏ꍇ�͌��ʂ�������܂��B
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		// TODO: �K�v�ɉ����ăG���[ �R�[�h��ύX���Ă��������B
		_tprintf(_T("�v���I�ȃG���[: MFC �̏��������ł��܂���ł����B\n"));
		nRetCode = 1;
	}
	else
	{
		// TODO: �A�v���P�[�V�����̓�����L�q����R�[�h�������ɑ}�����Ă��������B
	}

	return nRetCode;
}

ACCESSFROMVBA_API void WINAPI DoNothing()
{
    return;
}

//int�^�̒l���󂯎���āAint�^�̖߂�l��Ԃ�
ACCESSFROMVBA_API int WINAPI ReturnInt_Multi2(int i)
{
	return i * 2;
}

//int�^�̒l���󂯎���āA�����Œl��ύX���ĕԂ�
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
	
	std::wstring ws(L"GetStringByParam�ԋp�f�[�^");
	
	unsigned short vt = pvString->vt;

	VariantClear(pvString);

	if (vt == (VT_BSTR | VT_BYREF))
	{
		//VBA����A
		//GetStringByParam(String)
		//�ŌĂ΂ꂽ�ꍇ
		pvString->vt = vt;

		BSTR bs = SysAllocString(ws.c_str());
		INT iResult = SysReAllocString(pvString->pbstrVal, bs);
		SysFreeString(bs);
	}
	else if ((vt == VT_BSTR) || (vt == VT_EMPTY))
	{
		//VBA����A
		//GetStringByParam(Variant)
		//�ŌĂ΂ꂽ�ꍇ
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

    //�܂��J��
    SysFreeString(*pbstr);

    //�Ԃ�������istd::wstring�͎g�p���Ȃ��j
    std::string sReturn("GetStringByParamS �ԋp�f�[�^������");

    //BSTR����
    *pbstr = SysAllocStringByteLen(sReturn.c_str(), sReturn.length());

    return;
}

ACCESSFROMVBA_API BSTR WINAPI GetStringByRetValS()
{
    //�Ԃ�������istd::wstring�͎g�p���Ȃ��j
    std::string s("GetStringByRetValS �ԋp�f�[�^������");

    //BSTR����
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

    //VBA����̕�����́Achar*�Ŋi�[����Ă���
    //BSTR�̓r����'\0'�����鎖�͑z�肵�Ă��Ȃ��B
    std::string s((char*)sString);

    MessageBoxA(NULL, s.c_str(), "DLL", MB_OK | MB_ICONINFORMATION);

    return;
}