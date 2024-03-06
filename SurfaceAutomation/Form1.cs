using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Msword = Microsoft.Office.Interop.Word;
using Msexcel = Microsoft.Office.Interop.Excel;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Diagnostics;

namespace SurfaceAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            



        }

        

        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void bringToFront(string classame, string title)
        {
            // Get a handle to the Calculator application.
            IntPtr handle = FindWindow(classame, null);

            // Verify that Calculator is a running process.
            if (handle == IntPtr.Zero)
            {
                return;
            }

            // Make Calculator the foreground application
            SetForegroundWindow(handle);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bringToFront("ApplicationFrameWindow", null);
        }

        private void btnExcelCreated_Click(object sender, EventArgs e)
        {
            string fileName = "";
            bool IsCreated = false;
            var process = from p in System.Diagnostics.Process.GetProcessesByName("wfica32") select p;

            foreach (var p in process)
            {
                if(p.MainWindowTitle.Contains("Microsoft Excel - " + fileName + ".Prn - \\\\Remote"))
                {
                    IsCreated = true;
                }
            }

        }
        public void ConvertDtToCSV(string path, System.Data.DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            foreach(DataColumn col in dt.Columns)
            {
                sb.Append(col.ColumnName + ',');
            }
            sb.Remove(sb.Length - 1, 1);
            sb.Append(Environment.NewLine);

            foreach(DataRow row in dt.Rows)
            {
                for(int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(row[i].ToString() + ",");
                }
                sb.Append(Environment.NewLine);
            }
            System.IO.File.WriteAllText(path, sb.ToString());
        }
        public string SaveAsWordToXML(string fileName)
        {
            string xmlFileName = string.Empty;
            Msword._Application newApp = new Msword.Application();
            object source = fileName;
            object target = fileName.Substring(0, fileName.Length - 4) + "xml";
            xmlFileName = target.ToString();
            if(!System.IO.File.Exists(xmlFileName))
            {
                object unknown = Type.Missing;
                newApp.Documents.Open(ref source, ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown);
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXML;

                newApp.ActiveDocument.SaveAs2(ref target, ref format, ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown, ref unknown, ref unknown,
                    ref unknown, ref unknown);

                newApp.Documents.Close(true, Type.Missing, Type.Missing);
                newApp.Quit(ref unknown, ref unknown, ref unknown);
            }
            return xmlFileName;
        }
        public string ReadXmlFile(string xmlFileName)
        {
            StringBuilder sb = new StringBuilder();
            string result = new System.IO.StreamReader(xmlFileName).ReadToEnd();
            return result;
        }
        public void ReadPdf(string path)
        {
            string ss = string.Empty;
            string s = string.Empty;

            PdfReader pdr = new PdfReader(path);
            StringBuilder text = new StringBuilder();
            string docFilePath = path.Substring(0, path.Length - 3) + "docx";
            string xmlFilePath = path.Substring(0, path.Length - 4) + "xml";
            string xmlFileName = "";
            if(System.IO.File.Exists(xmlFilePath))
            {
                System.IO.File.Delete(xmlFilePath);
            }
            xmlFileName = SaveAsWordToXML(docFilePath);
            string input = ReadXmlFile(xmlFilePath).ToString();
            for (int i = 1; i<= pdr.NumberOfPages; i++)
            {
                ss = ss + PdfTextExtractor.GetTextFromPage(pdr, i);
                ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                s = s + PdfTextExtractor.GetTextFromPage(pdr, i, its);
            }
            string[] x = ss.Split('\n');
        }
        public void WordReader(string filename)
        {
            Msword.Application app = new Msword.Application();
            Msword.Document doc = new Msword.Document();
            object missing = System.Type.Missing;
            doc = app.Documents.Open(filename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            String read = String.Empty;
            List<string> data = new List<string>();
            foreach(Range rng in doc.StoryRanges)
            {
                data.Add(rng.Text);
            }
            ((_Document)doc).Close();
            ((_Application)app).Quit();

        }
        public void ConvertGrayImage(string path, string index)
        {
            Bitmap bm1 = new Bitmap(path);
            Bitmap bm2 = new Bitmap(bm1.Width, bm1.Height);
            for(int i = 0; i < bm1.Width; i++)
            {
                for(int j = 0; j < bm2.Height; j++)
                {
                    Color c = bm1.GetPixel(i, j);
                    int average = ((c.R + c.B + c.G) / 3);
                    if(average < 200)
                    {
                        bm2.SetPixel(i, j, Color.Black);
                    }
                    else
                    {
                        bm2.SetPixel(i, j, Color.White);
                    }
                }
            }
            bm2.Save(index);
            bm1.Dispose();
            bm2.Dispose();
        }
        public System.Data.DataTable XmlToDt(string path)
        {
            System.Data.DataTable dt = new System.Data.DataTable(); ;
            DataSet ds = new DataSet();
            ds.ReadXml(path);
            dt = ds.Tables[0];
            return dt;
        }
        public void ScreenCapture(string path)
        {
            Bitmap bm;
            Graphics g;
            bm = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            g = Graphics.FromImage(bm);
            g.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
            bm.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
            bm.Dispose();
        }
        public void changeResolution(int screenWidth, int screenHeight)
        {
            clsResolution.CResolution rs = new clsResolution.CResolution(screenWidth, screenHeight);
        }

        private void btnTxt2Csv_Click(object sender, EventArgs e)
        {
            //ConvertToXlsx(@"", @"")
        }

        void ConvertToXlsx(string sourcefile, string destfile)
        {
            int i, j;
            Msexcel.Application xlApp;
            Msexcel.Workbook xlWorkBook;
            Msexcel._Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            string[] lines, cells;
            lines = System.IO.File.ReadAllLines(sourcefile);
            xlApp = new Msexcel.Application();
            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Msexcel._Worksheet)xlWorkBook.ActiveSheet;
            for (i = 0; i < lines.Length; i++)
            {
                cells = lines[i].Split(new Char[] { '\t', ';' });
                for (j = 0; j < cells.Length; j++)
                    xlWorkSheet.Cells[i + 1, j + 1] = cells[j];
            }
            xlWorkBook.SaveAs(destfile, Msexcel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
        }
    }
}
