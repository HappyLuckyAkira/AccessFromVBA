// AccessFromVBA.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"
#include "AccessFromVBA.h"

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