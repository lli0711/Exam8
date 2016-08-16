Imports System.Runtime.InteropServices


Public Class Form1

    '********************************************************
    '*  LoadLibrary,FreeLibrary,GetProcAddress の宣言
    '********************************************************
    <DllImport("kernel32", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function LoadLibrary(ByVal lpFileName As String) As IntPtr
    End Function

    <DllImport("kernel32", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
    End Function

    <DllImport("kernel32", CharSet:=CharSet.Ansi, SetLastError:=True)> _
    Private Shared Function GetProcAddress(ByVal hModule As IntPtr, ByVal lpProcName As String) As IntPtr
    End Function

    '********************************************************
    '* ExtractTextの宣言
    '********************************************************
    <UnmanagedFunctionPointer(CallingConvention.Cdecl)> _
    Delegate Sub ExtractText( _
        <MarshalAs(UnmanagedType.BStr)> ByVal lpFileName As String, _
        ByVal bProp As Boolean, _
        <MarshalAs(UnmanagedType.BStr)> ByRef lpFileText As String _
        )

    ' Windows 64bit版でビルドするときはターゲットをx86にする必要があります
    ' プロパティ→コンパイル→詳細コンパイルオプション(A)→ターゲットCPU(U):AnyCPU→x86

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim ofd As New OpenFileDialog()

        ofd.FileName = ""
        ofd.InitialDirectory = ""
        ofd.Filter = "All Files(*.*)|*.*"
        ofd.FilterIndex = 2
        ofd.Title = "Open File.."
        ofd.RestoreDirectory = True
        ofd.CheckFileExists = True
        ofd.CheckPathExists = True

        If ofd.ShowDialog() = DialogResult.OK Then

            Dim fileText As String
            fileText = ""

            ' xd2txlib.dll をLoadLibrary経由でロード
            Dim TaragetDll As String = "xd2txlib.dll"
            Dim TaragetFunc As String = "ExtractText"
            Dim hModule As UInt32 = LoadLibrary(TaragetDll)
            If hModule = 0 Then
                TextBox1.Text = "xd2txlib.dll not found"
                Exit Sub
            End If

            ' 変換の実行
            Dim ptr As IntPtr
            ptr = GetProcAddress(hModule, "ExtractText")

            If ptr <> IntPtr.Zero Then
                Dim dllFunc As ExtractText = _
                Marshal.GetDelegateForFunctionPointer( _
                ptr, _
                GetType(ExtractText) _
                )

                Call dllFunc(ofd.FileName, False, fileText)

            End If

            Label1.Text = ofd.FileName
            TextBox1.Text = fileText

            ' DLLの解放
            Call FreeLibrary(hModule)

        End If


    End Sub
End Class
