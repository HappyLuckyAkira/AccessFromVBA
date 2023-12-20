// AccessFromVBA.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"
#include "AccessFromVBA.h"
#include <string>
#include <comutil.h>
#include <iostream>
#include <sstream>

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

std::wstring	convMbc2Wstr(const char* lpcszSrc);
std::wstring	convMbcBstr2Wstr(const BSTR& bstr);

ACCESSFROMVBA_API void WINAPI SetArrayGE(const LPSAFEARRAY* ppsa)
{
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult))
    {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    std::wostringstream ss;

    for (UINT i = 1; i <= uiDims; ++i)
    {
        LONG lLBound, lUBound;
        hResult = SafeArrayGetLBound(*ppsa, i, &lLBound);
        hResult = SafeArrayGetUBound(*ppsa, i, &lUBound);

        ss << i << "次元\n"
            << "　LBound：" << lLBound << "\n"
            << "　UBound：" << lUBound << "\n";
    }

    if (vt == VT_I4)
    {
        ss << L"データ型：Long\n";

        if (uiDims == 2)
        {
            LONG lIndex[2];

            LONG lLBound[2];
            LONG lUBound[2];

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound[0]);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound[0]);

            for (LONG i = lLBound[0]; i <= lUBound[0]; ++i)
            {
                hResult = SafeArrayGetLBound(*ppsa, 2, &lLBound[1]);
                hResult = SafeArrayGetUBound(*ppsa, 2, &lUBound[1]);

                lIndex[0] = i;

                for (LONG j = lLBound[1]; j <= lUBound[1]; ++j)
                {
                    lIndex[1] = j;

                    int iValue;
                    hResult = SafeArrayGetElement(*ppsa, lIndex, &iValue);

                    ss << iValue << L"\n";
                }
            }
        }
    }
    else if (vt == VT_BSTR)
    {
        ss << L"データ型：String\n";

        if (uiDims == 1)
        {
            LONG lLBound;
            LONG lUBound;

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound);

            for (LONG lIndex = lLBound; lIndex <= lUBound; ++lIndex)
            {
                BSTR bstr;

                SafeArrayGetElement(*ppsa, &lIndex, &bstr);

                ss << convMbcBstr2Wstr(bstr) << L"\n";

                SysFreeString(bstr);
            }
        }
    }
    else if (vt == VT_VARIANT)
    {
        ss << L"データ型：Variant\n";

        if (uiDims == 1)
        {
            LONG lLBound;
            LONG lUBound;

            hResult = SafeArrayGetLBound(*ppsa, 1, &lLBound);
            hResult = SafeArrayGetUBound(*ppsa, 1, &lUBound);

            for (LONG lIndex = lLBound; lIndex <= lUBound; ++lIndex)
            {
                VARIANT vValue;
                VariantInit(&vValue);
                hResult = SafeArrayGetElement(*ppsa, &lIndex, &vValue);
	   			wchar_t buffer[100]; // 適切なサイズに調整してください
                switch (vValue.vt)
                {
                case VT_I4:
					swprintf_s(buffer, L"%f", vValue.intVal);
					ss << buffer << L"\n";
                    break;
                case VT_R8:
					 swprintf_s(buffer, L"%f", vValue.dblVal);
                    ss << buffer << L"\n";
                    break;
                case VT_BSTR:
                    ss << vValue.bstrVal << L"\n";
                    break;
                case VT_BSTR | VT_BYREF:
                    ss << vValue.pbstrVal << L"\n";
                    break;
                default:
                    break;
                } ;

                VariantClear(&vValue);
            }
        }

    }
    else
    {
        MessageBox(NULL, L"No Data Type matched.", L"SetArrayGE", MB_OK | MB_ICONINFORMATION);

        return;
    }

    MessageBox(NULL, ss.str().c_str(), L"SetArrayGE", MB_OK | MB_ICONINFORMATION);

    return;
}

ACCESSFROMVBA_API void WINAPI SetArrayAD(const LPSAFEARRAY* ppsa)
{
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult))
    {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    std::wstringstream ss;

    for (UINT i = 1; i <= uiDims; ++i)
    {
        LONG lLBound, lUBound;
        hResult = SafeArrayGetLBound(*ppsa, i, &lLBound);
        hResult = SafeArrayGetUBound(*ppsa, i, &lUBound);

        ss << i << "次元\n"
            << "　LBound：" << lLBound << "\n"
            << "　UBound：" << lUBound << "\n";
    }

    if (vt == VT_I4)
    {
        ss << L"データ型：Long\n";

        //全要素数
        ULONG ulElems(1);
        for (ULONG i = 0; i < uiDims; ++i)
        {
            ulElems *= (*ppsa)->rgsabound[i].cElements;
        }

        int* piValue(0);

        hResult = SafeArrayAccessData(*ppsa, (void**)&piValue);

        for (ULONG i = 0; i < ulElems; ++i)
        {
//            ss << L"0x" << std::setfill(L'0') << std::right << std::setw(8) << std::hex << *piValue << L"\n";
        // 8桁の16進数表示（手動で0埋め）
			ss << L"0x";
			for (int j = 7; j >= 0; --j) {
				int nibble = (*piValue >> (4 * j)) & 0xF;
				ss << std::hex << nibble;
			}
			ss << L"\n";
            ++piValue;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    }
    else if (vt == VT_BSTR)
    {
        ss << L"データ型：String\n";

        BSTR* pbstr(0);

        hResult = SafeArrayAccessData(*ppsa, (void**)&pbstr);

        std::wstring ws = convMbcBstr2Wstr(*pbstr);

        for (ULONG i = 0; i < (*ppsa)->rgsabound[0].cElements; ++i)
        {
            ws = convMbcBstr2Wstr(*pbstr);
            ss << ws << L"\n";

            ++pbstr;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    }
    else if (vt == VT_VARIANT)
    {
        ss << L"データ型：Variant\n";

        VARIANT* pv(0);
        
        hResult = SafeArrayAccessData(*ppsa, (void**)&pv);

        for (ULONG i = 0; i < (*ppsa)->rgsabound[0].cElements; ++i)
        {
   			wchar_t buffer[100]; // 適切なサイズに調整してください

			switch (pv->vt)
            {
            case VT_I4:
				swprintf_s(buffer, L"%d", pv->intVal);
				ss << buffer << L"\n";
				//ss << std::to_wstring(pv->intVal) << L"\n";
                break;
            case VT_R8:
				swprintf_s(buffer, L"%f", pv->dblVal);
				ss << buffer << L"\n";
                //ss << std::to_wstring(pv->dblVal) << L"\n";
                break;
            case VT_BSTR:
                ss << convVstr2Wstr(*pv) << L"\n";
                break;
            default:
                break;
            }

            ++pv;
        }

        hResult = SafeArrayUnaccessData(*ppsa);
    }
    else
    {
        MessageBox(NULL, L"No Data Type matched.", L"SetArrayAD", MB_OK | MB_ICONINFORMATION);

        return;
    }

    MessageBox(NULL, ss.str().c_str(), L"SetArrayAD", MB_OK | MB_ICONINFORMATION);

    return;
}

ACCESSFROMVBA_API void WINAPI GetArrayPE(LPSAFEARRAY* ppsa)
{
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult))
    {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    LONG* plLBound = new LONG[uiDims];
    LONG* plUBound = new LONG[uiDims];

    for (UINT i = 0; i < uiDims; ++i)
    {
        SafeArrayGetLBound(*ppsa, i + 1, &plLBound[i]);
        SafeArrayGetUBound(*ppsa, i + 1, &plUBound[i]);
    }

    if (vt == VT_I4)
    {
        if (uiDims == 2)
        {
            LONG* plIndex = new LONG[uiDims];

            for (LONG i = plLBound[0]; i <= plUBound[0]; ++i)
            {
                plIndex[0] = i;

                for (LONG j = plLBound[1]; j <= plUBound[1]; ++j)
                {
                    plIndex[1] = j;

                    long lValue;
                    SafeArrayGetElement(*ppsa, plIndex, &lValue);
                    lValue *= 2;
                    SafeArrayPutElement(*ppsa, plIndex, &lValue);
                }
            }
        }
    }
    else if (vt == VT_R8)
    {
        if (uiDims == 1)
        {
            LONG lIndex;

            for (LONG i = plLBound[0]; i <= plUBound[0]; ++i)
            {
                lIndex = i;

                double dValue;
                SafeArrayGetElement(*ppsa, &lIndex, &dValue);
                dValue /= 2.0;
                SafeArrayPutElement(*ppsa, &lIndex, &dValue);
            }
        }
    }
    else if (vt == VT_VARIANT)
    {
        if (uiDims == 1)
        {
            LONG lIndex;

            for (LONG i = plLBound[0]; i <= plUBound[0]; ++i)
            {
                VARIANT vOrg;
                VariantInit(&vOrg);

                lIndex = i;
                SafeArrayGetElement(*ppsa, &lIndex, &vOrg);

                VARIANT vOut;
                VariantInit(&vOut);
                if (vOrg.vt == VT_EMPTY)
                {
                    vOut.vt = VT_BSTR;
                }
                else
                {
                    vOut.vt = vOrg.vt;
                }

                switch (vOrg.vt)
                {
                case VT_I4:
                    vOut.intVal = vOrg.intVal * 2;
                    SafeArrayPutElement(*ppsa, &lIndex, &vOut);
                    break;
                case VT_R8:
                    vOut.dblVal = vOrg.dblVal / 2.0;
                    SafeArrayPutElement(*ppsa, &lIndex, &vOut);
                    break;
                case VT_BSTR:
                {
					// Using swprintf_s instead of _itow
					wchar_t buffer[100]; // Adjust the buffer size accordingly
					swprintf_s(buffer, L"%ld", lIndex);
					std::wstring ws = L"GetArrayPE";
                    ws += buffer;
                    SysReAllocString(&vOut.bstrVal, ws.c_str());
                    SafeArrayPutElement(*ppsa, &lIndex, &vOut);
                    SysFreeString(vOut.bstrVal);
                }
                    break;
                default:
                    break;
                }

                VariantClear(&vOrg);
                VariantClear(&vOut);
            }
        }
    }
    else if (vt == VT_BSTR)
    {
        if (uiDims == 1)
        {
            LONG lIndex;

            for (LONG i = plLBound[0]; i <= plUBound[0]; ++i)
            {
                lIndex = i;

//                std::string sReturn = "GetArrayPE:";
//				sReturn += std::to_string(i + 1);
				char buffer[100]; // Adjust the buffer size accordingly
				// Using sprintf_s instead of std::to_string
				sprintf_s(buffer, "GetArrayPE:%d", i + 1);
				std::string sReturn(buffer);

                size_t lenByte = sReturn.length();

                BSTR bstr = SysAllocStringByteLen(sReturn.c_str(), lenByte);
                //&bstrではないので注意！！！(& は不要)
                SafeArrayPutElement(*ppsa, &lIndex, bstr);

                SysFreeString(bstr);
            }
        }
    }
    
    delete[] plLBound;
    delete[] plUBound;

    return;
}

ACCESSFROMVBA_API void WINAPI GetArrayAD(LPSAFEARRAY* ppsa)
{
    //格納されているデータ型の確認
    VARTYPE vt;
    HRESULT hResult = SafeArrayGetVartype(*ppsa, &vt);

    if (FAILED(hResult))
    {
        //VARIANTに格納できない（？）データの配列の場合、この判定ではNG
        return;
    }

    //次元数
    UINT uiDims = SafeArrayGetDim(*ppsa);

    //全要素数
    ULONG ulElems(1);
    for (ULONG i = 0; i < uiDims; ++i)
    {
        ulElems *= (*ppsa)->rgsabound[i].cElements;
    }

    if (vt == VT_I4)
    {
        int* piValue;
        SafeArrayAccessData(*ppsa, (void**)&piValue);
        for (ULONG i = 0; i < ulElems; ++i)
        {
            *piValue *= 3;

            ++piValue;
        }
        SafeArrayUnaccessData(*ppsa);
    }
    else if (vt == VT_BSTR)
    {
        BSTR* pbstr;
        SafeArrayAccessData(*ppsa, (void**)&pbstr);
        for (ULONG i = 0; i < ulElems; ++i)
        {
            //std::string ws("GetArrayAD");
            //ws += std::to_string(i);
            //UINT lenByte = ws.length();
			char buffer[50]; // Adjust the buffer size accordingly

			// Using sprintf_s instead of std::to_string
			sprintf_s(buffer, "GetArrayAD:%d", i);
			std::string ws(buffer);
			UINT lenByte = static_cast<UINT>(ws.length());

            SysFreeString(*pbstr);
            *pbstr = SysAllocStringByteLen(ws.c_str(), lenByte);

            ++pbstr;
        }
        SafeArrayUnaccessData(*ppsa);
    }
    else if (vt == VT_VARIANT)
    {
        VARIANT* pv;
        SafeArrayAccessData(*ppsa, (void**)&pv);
        for (ULONG i = 0; i < ulElems; ++i)
        {
            switch (pv->vt)
            {
            case VT_I4:
                pv->intVal *= 4;
                break;
            case VT_EMPTY:
            {
                //std::wstring ws(L"GetArrayAD_EMPTY");
                //ws += std::to_wstring(i);
				WCHAR buffer[50]; // Adjust the buffer size accordingly
				// Using swprintf_s instead of std::to_wstring
				swprintf_s(buffer, L"GetArrayAD_EMPTY%d", i);
				std::wstring ws(buffer);

                pv->vt = VT_BSTR;
                pv->bstrVal = SysAllocString(ws.c_str());
            }
                break;
            case VT_BSTR:
            {
//                std::wstring ws(L"GetArrayAD_BSTR");
//                ws += std::to_wstring(i);
				WCHAR buffer[50]; // Adjust the buffer size accordingly

				// Using swprintf_s to format the wide string
				swprintf_s(buffer, L"GetArrayAD_BSTR%d", i);

				std::wstring ws(buffer);
                SysReAllocString(&(pv->bstrVal), ws.c_str());
            }
                break;
            default:
                break;
            }

            ++pv;
        }
        SafeArrayUnaccessData(*ppsa);
    }

    return;
}

std::wstring    convMbc2Wstr(const char* lpcszSrc)
{
    if (!lpcszSrc)
    {
        return std::wstring(L"");
    }

    //必要なバッファサイズ取得

    /*
    cbMultiByte（第４引数）
    lpMultiByteStrパラメーターで示される文字列のサイズ（バイト単位）。 
    または、文字列がヌルで終了する場合、このパラメーターを-1に設定できます。 
    cbMultiByteが0の場合、関数は失敗することに注意してください。
    
    このパラメーターが-1の場合、関数は終端のヌル文字を含む入力文字列全体を処理します。 
    したがって、結果のUnicode文字列には終端のヌル文字が含まれ、
    関数によって返される長さにはこの文字が含まれます。

    このパラメーターが正の整数に設定されている場合、
    関数は指定されたバイト数を正確に処理します。 
    指定されたサイズに終端のヌル文字が含まれていない場合、
    結果のUnicode文字列はヌル終端されず、
    返される長さにはこの文字は含まれません。 
    */
    int iNeedBufferElems = MultiByteToWideChar(CP_ACP, 0, lpcszSrc, -1, NULL, 0);

    wchar_t* pwsz = new wchar_t[iNeedBufferElems];

    int iResult = MultiByteToWideChar(CP_ACP, 0, lpcszSrc, -1, pwsz, iNeedBufferElems);

    std::wstring ws(pwsz);

    delete[] pwsz;

    return ws;
}

std::wstring    convMbcBstr2Wstr(const BSTR& bstr)
{
    return convMbc2Wstr((const char*)bstr);
}

ACCESSFROMVBA_API void WINAPI GetArrayV(LPVARIANT pv)
{
    if (!(pv->vt & VT_ARRAY))
    {
        //配列でなければ処理終了
        return;
    }

    UINT uiDims = SafeArrayGetDim(pv->parray);

    if (uiDims != 1)
    {
        return;
    }

    HRESULT hResult;
    long lUbound;
    long lLbound;

    hResult = SafeArrayGetLBound(pv->parray, 1, &lLbound);
    hResult = SafeArrayGetUBound(pv->parray, 1, &lUbound);

    //データ型判定を簡略化するために、VT_ARRAYフラグをOFFする
    unsigned short vt = pv->vt ^ VT_ARRAY;

    if (vt == VT_I4)
    {
        for (long i = lLbound; i <= lUbound; ++i)
        {
            long lValue;
            hResult = SafeArrayGetElement(pv->parray, &i, &lValue);

            if (lValue != 0)
                lValue *= 2;
            else
                lValue = -1;

            hResult = SafeArrayPutElement(pv->parray, &i, &lValue);
        }
    }
    else if (vt == VT_BSTR)
    {
        for (long i = lLbound; i <= lUbound; ++i)
        {
            BSTR bstr = SysAllocString(L"");

            hResult = SafeArrayGetElement(pv->parray, &i, &bstr);

            std::wstring ws = std::wstring(bstr, SysStringLen(bstr));

            ws += L" GetArrayV ";
//            ws += std::to_wstring(i);
			// Convert integer to wide string using swprintf_s
			wchar_t buffer[20];  // Adjust the buffer size as needed
			swprintf_s(buffer, L"%ld", i);
			ws += buffer;

            SysReAllocString(&bstr, ws.c_str());

            hResult = SafeArrayPutElement(pv->parray, &i, bstr);

            SysFreeString(bstr);
        }
    }

    return;
}

ACCESSFROMVBA_API void WINAPI SetArrayV(const LPVARIANT pv)
{
    if (pv == NULL || !(pv->vt & VT_ARRAY))
    {
        //配列でなければ処理終了
        return;
    }

    std::wstringstream ss;

    UINT uiDims = SafeArrayGetDim(pv->parray);
    ss << L"次元数：" << uiDims << L"\n";

    HRESULT hResult;
    long* plUbound = new long[uiDims];
    long* plLbound = new long[uiDims];

    for (UINT i = 0; i < uiDims; ++i)
    {
        hResult = SafeArrayGetLBound(pv->parray, i + 1, &plLbound[i]);
        hResult = SafeArrayGetUBound(pv->parray, i + 1, &plUbound[i]);

        ss << i + 1 << L"次元：" << L"\n"
            << L"  LBound:" << plLbound[i] << L"\n"
            << L"  UBound:" << plUbound[i] << L"\n";
    }

    //データ型判定を簡略化するために、VT_ARRAYフラグをOFFする
    unsigned short vt = pv->vt ^ VT_ARRAY;
    
    ss << L"データ型：";

    if (vt == VT_I4)
    {
        ss << L"Long\n";

        if (uiDims == 2)
        {
            long lIndex[] = { 0, 0 };

            for (long i = plLbound[0]; i <= plUbound[0]; ++i)
            {
                lIndex[0] = i;

                for (long j = plLbound[1]; j <= plUbound[1]; ++j)
                {
                    lIndex[1] = j;
                    int iValue(0);
                    hResult = SafeArrayGetElement(pv->parray, lIndex, &iValue);

                    ss << iValue << L"\n";
                }
            }
        }
        else
        {
            long* plIndex = new long[uiDims];

            for (UINT i = 0; i <= uiDims; ++i)
            {
                plIndex[i] = plLbound[i];
            }

            for (long i = plLbound[0]; i <= plUbound[0]; ++i)
            {
                plIndex[0] = i;

                int iValue(0);
                hResult = SafeArrayGetElement(pv->parray, plIndex, &iValue);

                ss << iValue << L"\n";
            }

            delete[] plIndex;
        }
    }
    else if (vt == VT_BSTR)
    {
        ss << L"String\n";

        if (uiDims == 1)
        {
            for (long i = plLbound[0]; i <= plUbound[0]; ++i)
            {
                BSTR bstr = SysAllocString(L"");

                SafeArrayGetElement(pv->parray, &i, &bstr);

                std::wstring ws = std::wstring(bstr, SysStringLen(bstr));

                ss << ws.c_str() << L"\n";

                SysFreeString(bstr);
            }
        }
        else
        {
            ss << L"not 1 Dimensions.\n";
        }
    }
    else
    {
        ss << L"Other Data Type.\n";
    }

    delete[] plUbound;
    delete[] plLbound;

    MessageBox(NULL, ss.str().c_str(), L"SetArrayV", MB_OK | MB_ICONINFORMATION);

    return;
}