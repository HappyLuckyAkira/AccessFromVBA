// �ȉ��� ifdef �u���b�N�� DLL ����̃G�N�X�|�[�g��e�Ղɂ���}�N�����쐬���邽�߂� 
// ��ʓI�ȕ��@�ł��B���� DLL ���̂��ׂẴt�@�C���́A�R�}���h ���C���Œ�`���ꂽ ACCESSFROMVBA_EXPORTS
// �V���{���ŃR���p�C������܂��B���̃V���{���́A���� DLL ���g���v���W�F�N�g�Œ�`���邱�Ƃ͂ł��܂���B
// �\�[�X�t�@�C�������̃t�@�C�����܂�ł��鑼�̃v���W�F�N�g�́A 
// ACCESSFROMVBA_API �֐��� DLL ����C���|�[�g���ꂽ�ƌ��Ȃ��̂ɑ΂��A���� DLL �́A���̃}�N���Œ�`���ꂽ
// �V���{�����G�N�X�|�[�g���ꂽ�ƌ��Ȃ��܂��B
#ifdef ACCESSFROMVBA_EXPORTS
#define ACCESSFROMVBA_API __declspec(dllexport)
#else
#define ACCESSFROMVBA_API __declspec(dllimport)
#endif

// ���̃N���X�� AccessFromVBA.dll ����G�N�X�|�[�g����܂����B
class ACCESSFROMVBA_API CAccessFromVBA {
public:
	CAccessFromVBA(void);
	// TODO: ���\�b�h�������ɒǉ����Ă��������B
};

extern ACCESSFROMVBA_API int nAccessFromVBA;

ACCESSFROMVBA_API int fnAccessFromVBA(void);

extern "C" {
    ACCESSFROMVBA_API void WINAPI DoNothing();
	ACCESSFROMVBA_API int WINAPI ReturnInt_Multi2(int i);
	ACCESSFROMVBA_API void WINAPI ReturnIntByPointerParameter(int* pi);
	ACCESSFROMVBA_API void WINAPI SetString(VARIANT vString);	
	ACCESSFROMVBA_API void WINAPI GetStringByParam(VARIANT* pvString);
	ACCESSFROMVBA_API VARIANT WINAPI GetStringByRetVal();
	ACCESSFROMVBA_API void WINAPI GetStringByParamS(BSTR* pbstr);
	ACCESSFROMVBA_API BSTR WINAPI GetStringByRetValS();
	ACCESSFROMVBA_API void WINAPI SetStringS(const BSTR sString);
	ACCESSFROMVBA_API void WINAPI SetArrayGE(const LPSAFEARRAY* ppsa);
	ACCESSFROMVBA_API void WINAPI SetArrayAD(const LPSAFEARRAY* ppsa);
	ACCESSFROMVBA_API void WINAPI GetArrayPE(LPSAFEARRAY* ppsa);
	ACCESSFROMVBA_API void WINAPI GetArrayAD(LPSAFEARRAY* ppsa);
    ACCESSFROMVBA_API void WINAPI GetArrayV(LPVARIANT pv);
    ACCESSFROMVBA_API void WINAPI SetArrayV(const LPVARIANT pv);

	#pragma pack(4)
    struct SamplePack
    {
        short   nValue;
        double  dValue;
        long    lValue;
    };
	#pragma pack()

    ACCESSFROMVBA_API void WINAPI GetStructP(SamplePack* pst);
    ACCESSFROMVBA_API void WINAPI GetStructPArray(LPSAFEARRAY* ppsa);
    ACCESSFROMVBA_API void WINAPI SetStructP(const SamplePack* pst);
    ACCESSFROMVBA_API void WINAPI SetStructPArray(const LPSAFEARRAY* ppsa);

	struct WithStringStruct
	{
		VARIANT name;
		int string_length;
	};
    ACCESSFROMVBA_API void WINAPI SetStructWithString(const WithStringStruct* pStruct);
    ACCESSFROMVBA_API void WINAPI GetStructWithString(WithStringStruct* pStruct);

}