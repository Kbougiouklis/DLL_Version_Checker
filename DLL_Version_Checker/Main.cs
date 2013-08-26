using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;

using Microsoft.Win32;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Web;
using System.Drawing;



namespace DLL_Version_Checker
{

    public class Main
    {
        public static string MyCSSDefinition;
        public static string MyHTMLClosure;
        
        public string RootPath;
        public string FolderName;

        public string RTDescritpion;

        public void RootFolderChech(RichTextBox MyReachtextBox)
        {
            MyDefaultHtmlValues();
            FolderName = "ISMS Version 2.0";
            string Mypath = RootPath;
            CreateLog();
            CheckFolderContent(Mypath, MyReachtextBox);
            CheckSubFolderContent(Mypath, MyReachtextBox);
            UpdateLog();
        }

        public void CheckSubFolderContent(string path, RichTextBox MyReachtextBox)
        {
            string[] getsubfolders = Directory.GetDirectories(path);
            foreach (string subfolder in getsubfolders)
            {

                FileInfo fi = new FileInfo(subfolder);
                FolderName = fi.Name;
                DefineRichTextString(MyReachtextBox, fi.Name, 1, "");
                string MySubPath = path + "\\" + fi.Name;
                CheckFolderContent(MySubPath, MyReachtextBox);
                CheckSubFolderContent(MySubPath, MyReachtextBox);
            }

        }

        public void CheckFolderContent(string path, RichTextBox MyReachtextBox)
        {
            //string path = RootPath;// +"\\COMMON";
            string MyLogFile = Application.StartupPath + "\\DLL Versions.txt";

            FileStream _fs = new FileStream(MyLogFile, FileMode.Append, FileAccess.Write);
            StreamWriter _sw = new StreamWriter(_fs, System.Text.Encoding.Default);

            string[] array1 = Directory.GetFiles(path, "*.dll".ToLower());
            string[] array2 = Directory.GetFiles(path, "*.ocx".ToLower());
            string[] array3 = Directory.GetFiles(path, "*.exe".ToLower());

            try
            {
                _sw.WriteLine(@"<div class=""ex""> <FONT face=Tahoma color=royalblue size=2><STRONG>Folder Name: " + FolderName + @"</STRONG></FONT> <table border=""1"">

        <tr>
          <th>DLL File</th>
          <th>Version</th>");
            
                foreach (string name in array1)
                {
                    string filename = Path.GetFileName(name);
                    FileVersionInfo _fvi = FileVersionInfo.GetVersionInfo(name);
                    DefineRichTextString(MyReachtextBox, filename, 2, _fvi.FileVersion);
                    _sw.WriteLine("<tr><td>" + filename + "</td>" + "<td>" + _fvi.FileVersion + "</td>  </tr>");
                    //Console.WriteLine(name);
                }

                foreach (string name in array2)
                {
                    string filename = Path.GetFileName(name);
                    FileVersionInfo _fvi = FileVersionInfo.GetVersionInfo(name);
                    DefineRichTextString(MyReachtextBox, filename, 2, _fvi.FileVersion);
                    _sw.WriteLine("<tr><td>" + filename + "</td>" + "<td>" + _fvi.FileVersion + "</td>  </tr>");
                }

                foreach (string name in array3)
                {
                    string filename = Path.GetFileName(name);
                    FileVersionInfo _fvi = FileVersionInfo.GetVersionInfo(name);
                    DefineRichTextString(MyReachtextBox, filename, 2, _fvi.FileVersion);
                    _sw.WriteLine("<tr><td>" + filename + "</td>" + "<td>" + _fvi.FileVersion + "</td>  </tr>");
                }
                _sw.WriteLine("</table> \n </div>");
            }

            catch
            {

                MessageBox.Show("Could not update data");
            }

            _sw.Close();
            _fs.Close();

        }


        public void CreateLog()
        {

            string MyLogFile = Application.StartupPath + "\\DLL Versions.txt";

            FileStream _fs = new FileStream(MyLogFile, FileMode.Create, FileAccess.Write);
            StreamWriter _sw = new StreamWriter(_fs, System.Text.Encoding.Default);

            if (!File.Exists (MyLogFile))
                try
                {
                    _fs = File.Create(MyLogFile);
                    _sw.WriteLine("");

                }
              catch
                {

                     MessageBox.Show("Error Creating Log File");

                }
            else
                try
                {
                    _sw.WriteLine(MyCSSDefinition);
                }
                catch
                 {

                     MessageBox.Show("Error Creating Log File");
                 }
 
            _sw.Close();
            _fs.Close();

        }

        public void UpdateLog()
        {

            string MyLogFile = Application.StartupPath + "\\DLL Versions.txt";

            FileStream _fs = new FileStream(MyLogFile, FileMode.Append, FileAccess.Write);
            StreamWriter _sw = new StreamWriter(_fs, System.Text.Encoding.Default);

            if (!File.Exists(MyLogFile))
                try
                {
                    _fs = File.Create(MyLogFile);
                    _sw.WriteLine("");

                }
                catch
                {

                    MessageBox.Show("Error Creating Log File");

                }
            else
                try
                {
                    _sw.WriteLine(MyHTMLClosure);
                }
                catch
                {

                    MessageBox.Show("Error Creating Log File");
                }

            _sw.Close();
            _fs.Close();

            ChangeExtention();
        }

        public void ChangeExtention()
        {
            string MyLogFile = Application.StartupPath + "\\DLL Versions.txt";
            FileInfo file = new FileInfo(MyLogFile);
            MyHtmlDeletion();
            string newPath = Path.ChangeExtension(MyLogFile, ".html");
            file.MoveTo(newPath);
            MytxtDeletion();
            
        }

        static void MyHtmlDeletion()
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(Application.StartupPath + "\\DLL Versions.html");
            try
            {
                fi.Delete();
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine(e.Message);
            }
        }


        static void MytxtDeletion()
        {
            System.IO.FileInfo txtDel = new System.IO.FileInfo(Application.StartupPath + "\\DLL Versions.txt");
            try
            {
                txtDel.Delete();
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine(e.Message);
            }
        }


        public void MyDefaultHtmlValues()
        {
            #region MyCSSDefinition
            MyCSSDefinition = @"<html>
        <head>
        <style>

        h1
	        {
		        border-style:double;
		        width:650px;
		        background-color:#6495ed;
		        font-family: :""Times New Roman"";
		        font-variant:small-caps;
		        font-size: 1.5em;
		        text-align:center;
	        }

        div.ex
            {
                width:650px;
                padding:10px;
                border:2px solid gray;
                margin:0px;
            }

        table 
	        {
		        border:1px solid black;
		        width:650px;
	        }

        th
	        {
		        font-variant:small-caps;
		        border:1px solid black;
		        background-color: #6495ed;

	        }

        td
	        {
		        border:1px solid black;
		        'white-space:nowrap;
	        }

        </style>
        </head>

        <body>
        <FONT face=Tahoma color=royalblue size=3><STRONG>Ulysses Systems - Services Group</STRONG></FONT><BR/>
        <FONT face=Tahoma color=blue size=2><STRONG>Service Update Report: DLL file Check Results</STRONG></FONT><HR>"
        + @"<FONT face=Tahoma color=black size=2>Service package executed on machine " + System.Environment.MachineName + " by user " + System.Environment.UserName + " at " + DateTime.Now + " </FONT><BR>"
        + @"<FONT face=Tahoma color=black size=2></FONT>
        <FONT face=Tahoma color=black size=2></FONT><HR>";
            #endregion

            #region MyHTMLClosure
            MyHTMLClosure = @"
            </body>
            </html>";
            #endregion
        }


        public void DefineRichTextString(RichTextBox MyReachtextBox, string ObjectName, int ObjectValue, string FileVersion)
        {
            if (ObjectValue == 1)
            {
                MyReachtextBox.SelectionColor = Color.Green;
                MyReachtextBox.AppendText("\n" + "============================== "
                + "Content of Folder: " + ObjectName + " ==============================" + "\n" + "\n");
            }
            if (ObjectValue == 2)
            {
                MyReachtextBox.SelectionColor = Color.BlueViolet;
                MyReachtextBox.AppendText("FileName: " + ObjectName + "\t" + "\t" + "FileVersion: " + FileVersion + "\n");
            }
        }

    }



    public class MyErrorHandlers
    {
        public string MyErrortext;
        public string MyErrorHeaderMessage;
        public int MyErrorIdentifier;

        public void ErrorMessageDefinition()
        {
            switch (MyErrorIdentifier)
            {
                case 1:
                    MessageBox.Show(MyErrortext, MyErrorHeaderMessage, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                case 2:
                    MessageBox.Show(MyErrortext, MyErrorHeaderMessage, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
        }

    }
}
