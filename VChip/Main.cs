using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace VChip
{
    public partial class Main : Form
    {

        private string pathFileName;
        private OpenFileDialog openFileDialog;

        public Main()
        {
            InitializeComponent();
        }

        // Event Function

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            openDialog();
        }

        private void txtFile_TextChanged(object sender, EventArgs e)
        {
            //openDialog();
        }

        private void btnEncrypt_Click(object sender, EventArgs e)
        {
            if (isFileEmpty()) MessageBox.Show("Please fill all the fields.");
            else
            {
                this.Text = "VignereCipher - Encrypting...";
                //backgroundWorker1.RunWorkerAsync();
                converting("encrypt");
            }
        }

        private void btnDecrypt_Click(object sender, EventArgs e)
        {
            if (isFileEmpty()) MessageBox.Show("Please fill all the fields.");
            else
            {
                this.Text = "VignereCipher - Encrypting...";
                //backgroundWorker1.RunWorkerAsync();
                converting("encrypt");
            }
        }

        // Main Function

        private void openDialog()
        {
            openFileDialog = new OpenFileDialog();
            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                pathFileName = openFileDialog.FileName;
                txtFile.Text = pathFileName;
            }
        }

        private bool isFileEmpty()
        {
            string keyword = txtKeyword.Text;
            if (txtFile.TextLength == 0 || keyword.Trim().Length == 0) return true;
            else return false;
        }

        private bool IsAWord(string text)
        {
            //var regex = new Regex(@"\b[\w']+\b");
            var regex = new Regex(@"\b[a-zA-Z]+\b");
            var match = regex.Match(text);
            return match.Value.Equals(text);
        }

        private bool isKeyAlphanumeric(string key)
        {

            return false;
        }

        private void converting(string action)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = application.Documents.Open(pathFileName);
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            application.Visible = false;


            int wordCount = document.Words.Count;
            string keyword = txtKeyword.Text;


            for (int i = 1; i <= wordCount; i++)
            {
                //Range range = document.Content
                var w = document.Words[i];
                string s = w.Text.ToString();
                //range.Find.ClearFormatting();
                if (!string.IsNullOrWhiteSpace(s.Trim()) && IsAWord(s.Trim()))
                {

                    w.Select();
                    //Console.WriteLine(s);
                    if (action.Equals("encrypt")) application.Selection.TypeText(VignereCipher.Encipher(s, keyword));
                    else application.Selection.TypeText(VignereCipher.Decipher(s, keyword));
                }
            }

            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTimeString = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            Console.WriteLine("RunTime (total):" + elapsedTimeString);
            stopWatch.Reset();
        }

        
    }
}
