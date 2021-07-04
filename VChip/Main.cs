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
                this.Text = "VignereCipher - Decrypting...";
                //backgroundWorker1.RunWorkerAsync();
                converting("decrypt");
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

        private bool IsAlphanumeric(string str)
        {
            //var regex = new Regex(@"\b[\w']+\b");
            var regex = new Regex(@"\b[a-zA-Z]+\b");
            var match = regex.Match(str);
            return match.Value.Equals(str);
            
        }

        private void converting(string action)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = application.Documents.Open(pathFileName);

            VignereCipher vigenereCipher = new VignereCipher(txtKeyword.Text);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            application.Visible = false;


            int wordCount = document.Words.Count;
            string keyword = txtKeyword.Text;


            for (int i = 1; i <= wordCount; i++)
            {
                Range range = document.Content;
                var w = document.Words[i];
                string s = w.Text.ToString();
                range.Find.ClearFormatting();
                if (!string.IsNullOrWhiteSpace(s.Trim()) && IsAlphanumeric(s.Trim()))
                {

                    w.Select();
                    //Console.WriteLine(w.Text);
                    //if (action.Equals("encrypt")) application.Selection.TypeText(VignereCipher.Encipher(s, keyword));
                    if (action.Equals("encrypt"))
                    {
                        application.Selection.TypeText(vigenereCipher.Encrypt(s));
                        Console.WriteLine($"Next Key Position : {vigenereCipher.nextKeyUsedPosition}");
                    }
                    else
                    {
                        application.Selection.TypeText(vigenereCipher.Decrypt(s));
                        Console.WriteLine($"Next Key Position : {vigenereCipher.nextKeyUsedPosition}");
                    }
                }
            }

            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTimeString = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
            Console.WriteLine("RunTime (total):" + elapsedTimeString);
            stopWatch.Reset();
            vigenereCipher.Reset();
        }

        
    }
}
