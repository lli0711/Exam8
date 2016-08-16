' xd2txcom.dll (xdoc2txt com dll版)サンプル
'===============================================================================
' (1)xd2txcom.dllをregsvr32で登録してから実行してください。
'	regsvr32 xd2txcom.dll
' (2)64bit OSで実行するときは、%WINDIR%SysWOW64\CScript.exe で実行してください。
'	C:\Windows\SysWOW64\cscript.exe comtest.vbs
'===============================================================================

'クラスオブジェクトを取得する
Set obj = CreateObject("xd2txcom.Xdoc2txt.1")

' 文書のテキストを抽出する
Dim fileText
fileText = obj.ExtractText("sample.doc",False)

MsgBox fileText

