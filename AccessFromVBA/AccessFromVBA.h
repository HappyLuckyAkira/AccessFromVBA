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
}