using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// 動的ロード版

namespace Sample4
{
    public partial class Form1 : Form
    {
        // LoadLibrary、FreeLibrary、GetProcAddressをインポート
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr LoadLibrary(string lpFileName);
        [DllImport("kernel32", SetLastError = true)]
        internal static extern bool FreeLibrary(IntPtr hModule);
        [DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = false)]
        internal static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        // 関数をdelegateで宣言する
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int ExtractText(
          [MarshalAs(UnmanagedType.BStr)] String lpFileName,
          bool bProp,
          [MarshalAs(UnmanagedType.BStr)] ref String lpFileText);

        // Windows 64bit版でビルドするときはターゲットをx86にする必要があります
        //  プロパティ→ビルド→プラットフォームターゲット(G): AnyCPU→x86

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.FileName = "";
            ofd.InitialDirectory = "";
            ofd.Filter = "All Files(*.*)|*.*";
            ofd.FilterIndex = 2;
            ofd.Title = "Open File..";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                int l;

                string fileText = "";
                // 動的にdllをロードし、使用後に解放
                IntPtr handle = LoadLibrary("xd2txlib.dll");
                IntPtr funcPtr = GetProcAddress(handle, "ExtractText");

                ExtractText extractText = (ExtractText)Marshal.GetDelegateForFunctionPointer(funcPtr, typeof(ExtractText));
                l = extractText(ofd.FileName, false, ref fileText);
                FreeLibrary(handle);

                label1.Text = ofd.FileName;
                textBox1.Text = fileText;

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}