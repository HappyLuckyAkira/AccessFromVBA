// AccessFromVBA.cpp : DLL �A�v���P�[�V�����p�ɃG�N�X�|�[�g�����֐����`���܂��B
//

#include "stdafx.h"
#include "AccessFromVBA.h"

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