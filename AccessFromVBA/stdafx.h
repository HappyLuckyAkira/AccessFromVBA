// stdafx.h : 標準のシステム インクルード ファイルのインクルード ファイル、または
// 参照回数が多く、かつあまり変更されない、プロジェクト専用のインクルード ファイル
// を記述します。
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Windows ヘッダーから使用されていない部分を除外します。
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // 一部の CString コンストラクタは明示的です。

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // Windows ヘッダーから使用されていない部分を除外します。
#endif

#include <afx.h>
#include <afxwin.h>         // MFC のコアおよび標準コンポーネント
#include <afxext.h>         // MFC の拡張部分
#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC の Internet Explorer 4 コモン コントロール サポート
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>                     // MFC の Windows コモン コントロール サポート
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <iostream>
// Windows ヘッダー ファイル:
#include <windows.h>



// TODO: プログラムに必要な追加ヘッダーをここで参照してください。
