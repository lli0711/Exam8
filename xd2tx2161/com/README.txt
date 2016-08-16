□xd2txcom.dllについて

xdoc2txtのCOM DLL版です。
著作権・利用条件は、xdoc2txtに準じます。

□配布ファイル

xd2txcom.dll	xdoc2txt COM DLL版
comtest.vbs	xd2txcom.dll を呼び出すVBScriptサンプル

□定義

ProgID:	xd2txcom.Xdoc2txt.1

HRESULT ExtractText([in] BSTR lpFilePath, VARIANT_BOOL bProp, [out,retval] BSTR* lpFileText);

	[in]BSTR lpFilePath	入力ファイル名
	[in]VARIANT_BOOL bProp	True:Office文書のプロパティ表示 False:本文テキスト表示
	[out,retval]BSTR* lpFileText	抽出したテキスト(Unicode)

HRESULT ExtractTextEx([in] BSTR lpFilePath, VARIANT_BOOL bProp, [in] BSTR lpOptions,[out,retval] BSTR* lpFileText);

	[in]BSTR lpFilePath	入力ファイル名
	[in]VARIANT_BOOL bProp	True:Office文書のプロパティ表示 False:本文テキスト表示
	[in]BSTR lpOptions	コマンドラインオプション(-r -o -g -x のみ有効)
	[out,retval]BSTR* lpFileText	抽出したテキスト(Unicode)


□使用方法

1) VBScriptの例

Set obj = CreateObject("xd2txcom.Xdoc2txt.1")
fileText = obj.ExtractText("sample.doc",False)
MsgBox fileText


2) アドインで実行する例

(Excel2003の場合)
・ツール→アドイン→オートメーション
・Xdoc2txt Classを選択
・有効なアドインにチェックをつける
これでVBAから利用可能です。


□注意事項

コマンドプロンプトを「管理者として実行」し、Regsvr32でxd2txcom.dllをシステムに登録してください。
	regsvr32 xd2txcom.dll

xd2txcom.dllを呼び出すスクリプトを、64bit OSで実行するときは、%WINDIR%SysWOW64\CScript.exe で実行してください。
	C:\Windows\SysWOW64\cscript.exe comtest.vbs


