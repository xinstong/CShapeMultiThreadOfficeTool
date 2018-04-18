using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace PicToWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择Word和图片所在文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    this.buttonStartHandle.Enabled = false;
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                this.textBoxPath.Text = dialog.SelectedPath;
                this.buttonStartHandle.Enabled = true;
            }
        }

        delegate void ShowProgressDelegate(int totalStep, int currentStep);
        delegate void ShowHandleResultDelegate(int totalStep, int currentStep, string wordName, string picName, int result);
        delegate void RunTaskDelegate(string dirRootPath);
        private void buttonStartHandle_Click(object sender, EventArgs e)
        {
            RunTaskDelegate runTask = new RunTaskDelegate(RunTask);

            runTask.BeginInvoke(this.textBoxPath.Text, null, null);
        }

        void ShowProgress(int totalStep, int currentStep)
        {
            if( this.progressBarHandle.InvokeRequired)
            {
                ShowProgressDelegate showProgress = new ShowProgressDelegate(ShowProgress );

                this.BeginInvoke( showProgress, new object[] { totalStep, currentStep } );
            }
            else
            {
                this.progressBarHandle.Maximum = totalStep;
                this.progressBarHandle.Value = currentStep;
            }
        }

        private string item;
        void ShowHandleResult(int totalStep, int currentStep, string wordName, string picName, int result)
        {
            
            if (this.listBoxResult.InvokeRequired)
            {
                ShowHandleResultDelegate reportHandleResult = new ShowHandleResultDelegate(ShowHandleResult);

                this.BeginInvoke(reportHandleResult, new object[] { totalStep, currentStep, wordName, picName, result });
            }
            else
            {
                item = "";
                item += "(";
                item += currentStep;
                item += "/";
                item += totalStep;
                item += ")";
                if (result == 0)
                {
                    item += "OK.";
                }
                else
                {
                    item += "ERROR.";
                }
                item += "   [";
                item += wordName;
                item += "]   [";
                item += picName;
                item += "]";

                this.listBoxResult.Items.Add(item);
            }
        }

        void RunTask(string dirRootPath)
        {
            allFileListWord.Clear();
            allFileListPic.Clear();
            GetAllDirList(dirRootPath);

            for (int i = 0; i < allFileListWord.Count; i++)
            {
                int iResult = 0;
                do
                {
                    string wordName = allFileListWord[i].ToString();
                    string[] keyNameArr = wordName.Split('.');
                    if (keyNameArr.Length == 0)
                    {
                        iResult = -1;
                        break;
                    }
                    string keyName = keyNameArr[0];
                    bool bFinded = false;
                    string toInsertPicFullName = "";
                    foreach(object picName in  allFileListPic)
                    {
                        toInsertPicFullName = picName.ToString();
                        if (toInsertPicFullName.Contains(keyName))
                        {
                            bFinded = true;
                            break;
                        }
                    }
                    if (!bFinded)
                    {
                        iResult = -1;
                        break;
                    }

                    Document document = new Document(allFileListWord[i].ToString());
                    Section s = document.AddSection();
                    Paragraph p = s.AddParagraph();
                    DocPicture Pic = p.AppendPicture(Image.FromFile(toInsertPicFullName));
                    Pic.Width = 750;
                    Pic.Height = 468;

                    //Save and Launch
                    if (keyNameArr[keyNameArr.Length - 1].Equals(".docx"))
                    {
                        document.SaveToFile(wordName, FileFormat.Docx);
                    }
                    else
                    {
                        document.SaveToFile(wordName, FileFormat.Doc);
                    }

                    //System.Diagnostics.Process.Start(wordName);



                    ShowProgress(allFileListWord.Count, i + 1);
                }while(false);

                string wordShortName = "";
                string picShortName = "";
                wordShortName = allFileListWord[i].ToString();
                wordShortName = wordShortName.Substring(dirRootPath.Length, wordShortName.Length - dirRootPath.Length);
                picShortName = allFileListPic[i].ToString();
                picShortName = picShortName.Substring(dirRootPath.Length, picShortName.Length - dirRootPath.Length);
                ShowHandleResult(allFileListWord.Count, i + 1,
                    wordShortName, picShortName, iResult);
            }
        }

        private ArrayList allFileListWord = new ArrayList();
        private ArrayList allFileListPic = new ArrayList();
        
        public void GetAllDirList(string strBaseDir)
        {
            DirectoryInfo di = new DirectoryInfo(strBaseDir);

            foreach(DirectoryInfo subDi in di.GetDirectories())
            {
                GetAllDirList(subDi.FullName);
            }

            foreach (FileInfo file in di.GetFiles("*.doc"))
            {
                allFileListWord.Add(file.FullName);
            }
            foreach (FileInfo file in di.GetFiles("*.docx"))
            {
                allFileListWord.Add(file.FullName);
            }

            foreach (FileInfo file in di.GetFiles("*.jpg"))
            {
                allFileListPic.Add(file.FullName);
            }
            foreach (FileInfo file in di.GetFiles("*.png"))
            {
                allFileListPic.Add(file.FullName);
            }
        }

        private void textBoxPath_TextChanged(object sender, EventArgs e)
        {
            this.buttonStartHandle.Enabled = Directory.Exists(this.textBoxPath.Text);
        }

        private void listBoxResult_DrawItem(object sender, DrawItemEventArgs e)
        {

        }
    }
}
