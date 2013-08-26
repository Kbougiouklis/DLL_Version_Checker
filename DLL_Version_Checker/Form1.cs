using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;

namespace DLL_Version_Checker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string MyLocalTAPath;
        Main _mn = new Main();
        MyErrorHandlers _eh = new MyErrorHandlers();

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "This tool will check the content of the folder called ISMS Version 2.0";
            GetregistryValues();
            label1.ForeColor = Color.Blue;
            label1.Text = "Application Path: " + MyLocalTAPath;
            richTextBox1.Clear();
            richTextBox1.ReadOnly = true;
            button1.Text = "Start";
            this.MaximumSize = new Size(637, Screen.PrimaryScreen.Bounds.Height);
            EndRichTextBox();
        }

        public void InitialRichTextBoxValue()
        {
            richTextBox1.Clear();
            richTextBox1.SelectionColor = Color.Green;
            richTextBox1.AppendText("\n" + "============================== "
            + " Content of Folder: ISMS Version 2.0 ==============================" + "\n" + "\n");
        }

        public void EndRichTextBox()
        {
            richTextBox1.SelectionStart = 0;
        }

        public void GetregistryValues()
        {
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("ISM-Solutions");
            //RegistryKey rg = regKey.OpenSubKey("SOFTWARE\\ISM-Solutions\\Task Assistant\\Me");
            if (regKey == null)
            {
                _eh.MyErrortext = @"Registry Key with Name " + @"""ISM-Solutions""" + @" does not exist";
                _eh.MyErrorHeaderMessage = "Registry Error";
                _eh.MyErrorIdentifier = 1;
                _eh.ErrorMessageDefinition();
                this.Close();
            }

            else
            {
                RegistryKey rg = regKey.OpenSubKey("Task Assistant\\Local_Config");
                MyLocalTAPath = (string)rg.GetValue("AppPath".ToUpper());
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            InitialRichTextBoxValue();
            _mn.RootPath = MyLocalTAPath;
            _mn.RootFolderChech(richTextBox1);
            _eh.MyErrortext = "Operation is now complete!" + "\nAn HTML Report has been created on the following directory:" + "\n" + Application.StartupPath;
            _eh.MyErrorHeaderMessage = "Information";
            _eh.MyErrorIdentifier = 2;
            _eh.ErrorMessageDefinition();
        }
    }
}
