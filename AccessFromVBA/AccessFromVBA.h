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
}