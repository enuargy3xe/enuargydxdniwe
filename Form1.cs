using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace WOW
{
    public partial class Form1 : Form
    {
        Word._Application oWord = new Word.Application();

        public Form1()
        {
            InitializeComponent();
        }

        private Word._Document GetDoc(string path)
        {
            Word._Document oDoc = oWord.Documents.Add(path);
            SetBookmarks(oDoc);
            return oDoc;
        }

        private void SetBookmarks(Word._Document oDoc)
        {
            oDoc.Bookmarks["prodTitle"].Range.Text = textBox1.Text;
            oDoc.Bookmarks["prodPrice"].Range.Text = textBox2.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();
            string savePath = Environment.CurrentDirectory + "\\" + DateTime.Now.Date.ToShortDateString() +"_" +rnd.Next(1,1001).ToString()+ "_ToPrint.docx";

            Word._Document oDoc = GetDoc(Environment.CurrentDirectory + "\\check.docx");
            oDoc.SaveAs(FileName:savePath);
            oDoc.Close();
            Process.Start(savePath);
        }
    }
}
