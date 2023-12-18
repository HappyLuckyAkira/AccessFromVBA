// 以下の ifdef ブロックは DLL からのエクスポートを容易にするマクロを作成するための 
// 一般的な方法です。この DLL 内のすべてのファイルは、コマンド ラインで定義された ACCESSFROMVBA_EXPORTS
// シンボルでコンパイルされます。このシンボルは、この DLL を使うプロジェクトで定義することはできません。
// ソースファイルがこのファイルを含んでいる他のプロジェクトは、 
// ACCESSFROMVBA_API 関数を DLL からインポートされたと見なすのに対し、この DLL は、このマクロで定義された
// シンボルをエクスポートされたと見なします。
#ifdef ACCESSFROMVBA_EXPORTS
#define ACCESSFROMVBA_API __declspec(dllexport)
#else
#define ACCESSFROMVBA_API __declspec(dllimport)
#endif

// このクラスは AccessFromVBA.dll からエクスポートされました。
class ACCESSFROMVBA_API CAccessFromVBA {
public:
	CAccessFromVBA(void);
	// TODO: メソッドをここに追加してください。
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
}