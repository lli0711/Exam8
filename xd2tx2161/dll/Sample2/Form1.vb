Imports System.Runtime.InteropServices

Public Class Form1

    'Declare Unicode Sub ExtractText Lib "xd2txlib.dll" ( _
    '     <MarshalAs(UnmanagedType.BStr)> ByVal lpFileName As String, _
    '     ByVal bProp As Boolean, _
    '     <MarshalAs(UnmanagedType.BStr)> ByRef lpFileText As String _
    '     )

    ' Windows 64bit版でビルドするときはターゲットをx86にする必要があります
    ' プロパティ→コンパイル→詳細コンパイルオプション(A)→ターゲットCPU(U):AnyCPU→x86

    <DllImport("xd2txlib.dll", SetLastError:=True, CharSet:=CharSet.Unicode, CallingConvention:=CallingConvention.Cdecl)> _
    Public Shared Function _
        ExtractText( _
        <MarshalAs(UnmanagedType.BStr)> ByVal lpFileName As String, _
         ByVal bProp As Boolean, _
         <MarshalAs(UnmanagedType.BStr)> ByRef lpFileText As String _
        ) As Integer
    End Function

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

            ExtractText(ofd.FileName, False, fileText)

            Label1.Text = ofd.FileName
            TextBox1.Text = fileText

        End If






    End Sub
End Class
