' xd2txcom.dll (xdoc2txt com dll��)�T���v��
'===============================================================================
' (1)xd2txcom.dll��regsvr32�œo�^���Ă�����s���Ă��������B
'	regsvr32 xd2txcom.dll
' (2)64bit OS�Ŏ��s����Ƃ��́A%WINDIR%SysWOW64\CScript.exe �Ŏ��s���Ă��������B
'	C:\Windows\SysWOW64\cscript.exe comtest.vbs
'===============================================================================

'�N���X�I�u�W�F�N�g���擾����
Set obj = CreateObject("xd2txcom.Xdoc2txt.1")

' �����̃e�L�X�g�𒊏o����
Dim fileText
fileText = obj.ExtractText("sample.doc",False)

MsgBox fileText

