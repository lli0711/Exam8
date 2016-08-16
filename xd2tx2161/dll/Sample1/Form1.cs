using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace Sample4
{
    public partial class Form1 : Form
    {
        [DllImport("xd2txlib.dll", CharSet = CharSet.Unicode,
            CallingConvention = CallingConvention.Cdecl)]
      public static extern int ExtractText(
          [MarshalAs(UnmanagedType.BStr)] String lpFileName,
          bool bProp,
          [MarshalAs(UnmanagedType.BStr)] ref String lpFileText );

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

                string fileText = "";
                int l = ExtractText(ofd.FileName, false, ref fileText);

                label1.Text = ofd.FileName;
                textBox1.Text   =  fileText;

            }
        }
    }
}