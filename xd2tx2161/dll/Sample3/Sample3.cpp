// Sample3.cpp : コンソール アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"
#include "OleAuto.h"
#include "comutil.h"
#include <io.h>
#include <fcntl.h>



typedef int (*FUNCTYPE)(BSTR lpFilePath, bool bProp,  BSTR*lpFileText);

int _tmain(int argc, _TCHAR* argv[])
{

	HINSTANCE   hInstDLL;       /*  DLLのインスタンスハンドル                */
	FUNCTYPE       ExtractText;    /*  DLLの関数へのポインタ    */

	if((hInstDLL=LoadLibrary( _T("xd2txlib.dll") ) ) == NULL ) {
		abort();
	}

	ExtractText = (FUNCTYPE)GetProcAddress(hInstDLL,"ExtractText");
	if( ExtractText != NULL ) {

		BSTR fileText = ::SysAllocString( _T("") );

		int nFileLength = ExtractText( _T("sample.doc"), false, &fileText );

		BYTE BOM[2] = {0xff, 0xfe};
		_setmode( _fileno( stdout ), _O_BINARY );
		fwrite( BOM, 2, 1, stdout );
		fwrite( (LPCTSTR)fileText, nFileLength, 1, stdout );


    }
	return 0;
}

