��xd2txcom.dll�ɂ���

xdoc2txt��COM DLL�łł��B
���쌠�E���p�����́Axdoc2txt�ɏ����܂��B

���z�z�t�@�C��

xd2txcom.dll	xdoc2txt COM DLL��
comtest.vbs	xd2txcom.dll ���Ăяo��VBScript�T���v��

����`

ProgID:	xd2txcom.Xdoc2txt.1

HRESULT ExtractText([in] BSTR lpFilePath, VARIANT_BOOL bProp, [out,retval] BSTR* lpFileText);

	[in]BSTR lpFilePath	���̓t�@�C����
	[in]VARIANT_BOOL bProp	True:Office�����̃v���p�e�B�\�� False:�{���e�L�X�g�\��
	[out,retval]BSTR* lpFileText	���o�����e�L�X�g(Unicode)

HRESULT ExtractTextEx([in] BSTR lpFilePath, VARIANT_BOOL bProp, [in] BSTR lpOptions,[out,retval] BSTR* lpFileText);

	[in]BSTR lpFilePath	���̓t�@�C����
	[in]VARIANT_BOOL bProp	True:Office�����̃v���p�e�B�\�� False:�{���e�L�X�g�\��
	[in]BSTR lpOptions	�R�}���h���C���I�v�V����(-r -o -g -x �̂ݗL��)
	[out,retval]BSTR* lpFileText	���o�����e�L�X�g(Unicode)


���g�p���@

1) VBScript�̗�

Set obj = CreateObject("xd2txcom.Xdoc2txt.1")
fileText = obj.ExtractText("sample.doc",False)
MsgBox fileText


2) �A�h�C���Ŏ��s�����

(Excel2003�̏ꍇ)
�E�c�[�����A�h�C�����I�[�g���[�V����
�EXdoc2txt Class��I��
�E�L���ȃA�h�C���Ƀ`�F�b�N������
�����VBA���痘�p�\�ł��B


�����ӎ���

�R�}���h�v�����v�g���u�Ǘ��҂Ƃ��Ď��s�v���ARegsvr32��xd2txcom.dll���V�X�e���ɓo�^���Ă��������B
	regsvr32 xd2txcom.dll

xd2txcom.dll���Ăяo���X�N���v�g���A64bit OS�Ŏ��s����Ƃ��́A%WINDIR%SysWOW64\CScript.exe �Ŏ��s���Ă��������B
	C:\Windows\SysWOW64\cscript.exe comtest.vbs


