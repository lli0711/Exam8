using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

/**
 *  xdoc2txt dll呼び出しサンプル
 *  実行にはxd2txlib.dllが必要です
 *  指定したフォルダ以下のファイルを検索し、ファイル名とファイル先頭部分をログに記録します。
 *  実行終了後、ログを表示します。
 *  xdoc2txtが対応している拡張子以外はスキップします。
 *  ファイルサイズ上限は32MBに設定しています。
 */

namespace FileFind
{
    public partial class Form1 : Form
    {
        const long MAXFILELEN = 32 * 1024 * 1024;  // ファイルサイズ上限 32MBに設定

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

        string[] extList = {
            ".doc", //  MS-Word / OASYS2文書(分離型)
            ".wbk", // MS-Word backup
            ".dot", // MS-Word テンプレート
            ".rtf", // リッチテキスト
            ".wri", // Write
            ".oas", // OASYS
            ".oa2", // OASYS
            ".oa3", // OASYS
            ".xls", // EXCEL
            ".xlk",	// EXCEL
            ".xlt", // EXCEL
            ".pdf", // PDF
            ".html",
            ".htm",
            ".shtml", 
            ".mht",
            ".xml",
            ".xhtml", 
            ".sgml",
            ".wbn",
            ".docx",    // Microsoft Office Word 2007(Office Open XML)
            ".docm",    // Microsoft Office Word 2007 with macro(Office Open XML)
            ".xlsx",    // Microsoft Office Excel 2007(Office Open XML)
            ".xlsm",    // Microsoft Office Excel 2007 with macro(Office Open XML)
            ".pptx",    // Microsoft Office PowerPoint 2007(Office Open XML)
            ".pptm",    // Microsoft Office PowerPoint with macro 2007(Office Open XML)
            ".sxw",		// OpenOffice.org Writer
            ".sxc",		// OpenOffice.org Calc
            ".sxi",		// OpenOffice.org Impress
            ".sxd",		// OpenOffice.org Draw
            ".odt",		// Open Document Format(ODF) Writer
            ".ods",		// Open Document Format(ODF) Sheet
            ".odp",		// Open Document Format(ODF) Presentation
            ".odg",		// Open Document Format(ODF) Draw
            ".jtt",	    // Ver.8/9
            ".jvw",	    // Ver.7
            ".juw",	    // Ver.6
            ".jtw",	    // Ver.5
            ".ppt",	    // パワーポイント
            ".wj2",	    // DOS版R2.2J、R2.3J
            ".wj3",	    // DOS版R2.4J
            ".wk1",	    // US版のDOS版R2
            ".wk3",	    // US版のWindows版R1
            ".wk4",	    // Windows版R5J
            ".123",	    // Windows版97
            ".12m",	    // Windows版97
            ".bun",	    // 新松
            ".txt"  //  テキストファイル
            };

        // バックグラウンド処理用
        private BackgroundWorker g_BackgroundWorker;
        private StreamWriter writer;

        private int g_files = 0;    //  処理対象となったファイル個数

        /// <summary>
        /// 
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            // BackgroundWorker
            g_BackgroundWorker = new BackgroundWorker();
            g_BackgroundWorker.DoWork += new DoWorkEventHandler(FileFindAsync);


        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
        }


        /// <summary>
        /// バックグラウンドでファイル検索
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FileFindAsync(object sender, DoWorkEventArgs e)
        {
            Object[] args = e.Argument as Object[];
            string path = args[0] as string;

            g_files = 0;
            textBox2.Text = textBox2.Text  + DateTime.Now.ToString() + " - start\r\n";

            //  ファイルにログを作成
            string logFileName = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal) + "\\findfile.log";

            Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
            writer = new StreamWriter(logFileName, false, sjisEnc);
            writer.WriteLine(DateTime.Now.ToString() + " - start");

            //  指定したパス以下のファイルを検索
            findFiles(path);

            //  処理終了を記録
            textBox2.Text = textBox2.Text + ((Int32)g_files).ToString()+ " files \r\n";
            textBox2.Text = textBox2.Text + DateTime.Now.ToString() + " - end\r\n";

            writer.WriteLine( ((Int32)g_files).ToString()+ " files" );
            writer.WriteLine(DateTime.Now.ToString() + " - end");
            writer.Close();

            //  ログファイルを開く
            System.Diagnostics.Process p = System.Diagnostics.Process.Start(logFileName);
        }

        /// <summary>
        /// 指定したディレクトリ以下のファイルを検索する
        /// </summary>
        /// <param name="dir"></param>
        public void findFiles(string dir)
        {

            string[] files = Directory.GetFiles(dir);
            foreach (string path in files)
            {
                FileInfo finfo = new FileInfo(path);
                long fileSize = finfo.Length;

                //  ファイルサイズが上限を超えるものはスキップ
                if (fileSize > MAXFILELEN) continue;

                //  拡張子を取得
                string ext = Path.GetExtension(path.ToLower());
                //  xdoc2txtのサポートする拡張子なら内容を取得
                if ( Array.IndexOf(extList,ext) > -1 ) 
                {
                    textBox3.Text = path;
                    string fileText = xdoc2txt(path);

                    //  パス名とファイル先頭をログに記録
                    writer.WriteLine("======================================");
                    writer.WriteLine(path);
                    fileText = fileText.Replace("\r\n", "/");
                    writer.WriteLine(fileText.Substring(0,Math.Min(80,fileText.Length)));

                    ++g_files;
                }
            }

            string[] dirs = Directory.GetDirectories(dir);
            foreach (string s in dirs)
            {
                findFiles(s);
            }
        }

        /// <summary>
        /// xdoc2txt DLLでテキストを抽出する
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private string xdoc2txt(string s)
        {

            string fileText = "";

            // 動的にdllをロードし、使用後に解放
            IntPtr handle = LoadLibrary("xd2txlib.dll");
            IntPtr funcPtr = GetProcAddress(handle, "ExtractText");

            ExtractText extractText = (ExtractText)Marshal.GetDelegateForFunctionPointer(funcPtr, typeof(ExtractText));
            int textlen = extractText(s, false, ref fileText);

            FreeLibrary(handle);

            return fileText;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //上部に表示する説明テキストを指定する
            fbd.Description = "フォルダを指定してください。";
            //ルートフォルダを指定する
            //デフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = "";
            fbd.ShowNewFolderButton = false;

            //ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                //選択されたフォルダを表示する
                textBox1.Text = fbd.SelectedPath;
            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void findButton_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                textBox2.Text = "";

                g_BackgroundWorker.RunWorkerAsync(
                    new Object[]{ textBox1.Text }
                );


            }
        }

    }
}
