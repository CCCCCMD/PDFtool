using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
//using PdfSharp.Pdf.IO;
//using PdfSharp.Pdf;
//using PdfSharp.Drawing;
using Document = iTextSharp.text.Document;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace pdf
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        private static List<FileInfo> lst = new List<FileInfo>();
        private String inputPath;
        private String outputPath;

        private static void getdir(string path, List<FileInfo> list)
        {

            string[] dir = Directory.GetDirectories(path); //文件夹列表   
            DirectoryInfo fdir = new DirectoryInfo(path);
            FileInfo[] file = fdir.GetFiles();
            if (file.Length != 0 || dir.Length != 0) //当前目录文件或文件夹不为空                   
            {
                foreach (FileInfo f in file) //显示当前目录所有文件 
                    list.Add(f);
                foreach (string d in dir)
                {
                    getdir(d, list);//递归   
                }
            }
        }

        private bool WordToPDF(string inePath, string outputPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Word.Document document = null;
            try
            {
                application.Visible = false;
                document = application.Documents.Open(inePath);
                document.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
                result = true;
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                result = false;
            }
            finally
            {
                document.Close();
                application.Quit();
            }
            return result;
        }

        private bool ExcelToPDF(string sourcePath, string targetPath)
        {
            bool result;
            object missing = Type.Missing;
            Excel.ApplicationClass application = null;
            Excel.Workbook workBook = null;
            try
            {
                application = new Excel.ApplicationClass();
                object target = targetPath;
                object type = Excel.XlFixedFormatType.xlTypePDF;
                workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);

                workBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;
            }
            catch
            {
                result = false;
            }
            finally
            {
                    workBook.Close(true, missing, missing);
                    application.Quit();
            }
            return result;
        }

        //private void comPDF(String filename1, String filename2,String out3)
        //{
        //    PdfDocument inputDocument1 = PdfReader.Open(filename1, PdfDocumentOpenMode.Import);
        //    PdfDocument inputDocument2 = PdfReader.Open(filename2, PdfDocumentOpenMode.Import);
        //    PdfDocument outputDocument = new PdfDocument();
        //    // Show consecutive pages facing. Requires Acrobat 5 or higher.
        //    outputDocument.PageLayout = PdfPageLayout.TwoColumnLeft;
        //    XFont font = new XFont("Verdana", 10, XFontStyle.Bold);
        //    XStringFormat format = new XStringFormat();
        //    format.Alignment = XStringAlignment.Center;
        //    format.LineAlignment = XLineAlignment.Far;
        //    //XGraphics gfx;
        //    //XRect box;
        //    int count = Math.Max(inputDocument1.PageCount, inputDocument2.PageCount);
        //    for (int idx = 0; idx < count; idx++)
        //    {
        //        PdfPage page1 = inputDocument1.PageCount > idx ?
        //          inputDocument1.Pages[idx] : new PdfPage();
        //        PdfPage page2 = inputDocument2.PageCount > idx ?
        //          inputDocument2.Pages[idx] : new PdfPage();

        //        // Add both pages to the output document
        //        page1 = outputDocument.AddPage(page1);
        //        page2 = outputDocument.AddPage(page2);

        //        // Write document file name and page number on each page
        //        //gfx = XGraphics.FromPdfPage(page1);
        //        //box = page1.MediaBox.ToXRect();
        //        //box.Inflate(0, -10);
        //        //gfx.DrawString(String.Format("{0} • {1}", filename1, idx + 1), font, XBrushes.Red, box, format);
        //        //gfx = XGraphics.FromPdfPage(page2);
        //        //box = page2.MediaBox.ToXRect();
        //        //box.Inflate(0, -10);
        //        //gfx.DrawString(String.Format("{0} • {1}", filename2, idx + 1), font, XBrushes.Red, box, format);
        //    }
        //    string filename = out3;
        //    outputDocument.Save(filename);
        //}

        private void comPDF(List<string> filelist, string path0)
        {
            PdfReader reader;
            List<PdfReader> readerList = new List<PdfReader>();
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(path0, FileMode.Create));
            document.Open();
            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage newPage;
            for (int i = 0; i < filelist.Count; i++)
            {
                reader = new PdfReader(filelist[i]);
                int iPageNum = reader.NumberOfPages;
                for (int j = 1; j <= iPageNum; j++)
                {
                    document.NewPage();
                    newPage = writer.GetImportedPage(reader, j);
                    cb.AddTemplate(newPage, 0, 0);
                }
                readerList.Add(reader);
            }
            document.Close();
            foreach (var rd in readerList)//清理占用
            {
                rd.Dispose();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            lst.Clear();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                inputPath = dialog.SelectedPath;
                textBox1.Text = inputPath;
                getdir(inputPath, lst);
                foreach (FileInfo f in lst)
                {
                    String type = "其他文档";
                    if (f.Extension == ".docx")
                        type = "Word文档";
                    if (f.Extension == ".xlsx")
                        type = "Execel文档";
                    if (f.Extension == ".txt")
                        type = "文本文档";
                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = index + 1;
                    this.dataGridView1.Rows[index].Cells[1].Value = f.Name;
                    this.dataGridView1.Rows[index].Cells[2].Value = type;
                }
                progressBar1.Value = 0;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = lst.Count;
                label3.Text = "已载入";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                outputPath = dialog.SelectedPath;
                textBox2.Text = outputPath;
                int i = outputPath.Length;
                if (outputPath.Length == 3)
                    outputPath = outputPath.Remove(2);
            }
        }

        private String toPDF(FileInfo f, String s, int flag)
        {
            String newPath = outputPath;
            String newName = f.Name;
            newName = newName.Replace(s, "pdf");
            if (flag == 1)
            {
                newPath = f.DirectoryName;
                newPath = newPath.Replace(inputPath, outputPath);
                if (!Directory.Exists(newPath))
                    Directory.CreateDirectory(newPath);
                newPath = newPath + "\\" + newName;
            }
            else
            {
                newPath = newPath + "\\temp\\";
                Directory.CreateDirectory(newPath);
                newPath = newPath + newName;
            }
            return newPath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int flag = 1;
            if (radioButton2.Checked)
            {
                flag = 0;
                progressBar1.Maximum++;
            }
            foreach (FileInfo f in lst)
            {
                if (f.Extension == ".docx")
                    WordToPDF(f.FullName, toPDF(f, "docx", flag));
                if (f.Extension == ".txt")
                    WordToPDF(f.FullName, toPDF(f, "txt", flag));
                if (f.Extension == ".xlsx")
                    ExcelToPDF(f.FullName, toPDF(f, "xlsx", flag));
                progressBar1.Value++;
                label3.Text = progressBar1.Value + "/" + lst.Count;
            }
            if (radioButton2.Checked)
            {
                List<FileInfo> l = new List<FileInfo>();
                List<String> pdf = new List<String>();
                DirectoryInfo dir = new DirectoryInfo(outputPath);
                getdir(outputPath + "\\temp", l);
                foreach (FileInfo f in l)
                    pdf.Add(f.FullName);
                string na = dir.Name;
                label3.Text = "正在合并";
                comPDF(pdf,outputPath+"\\"+dir.Name+".pdf");

                //String path1 = "", path2, outputpath_t, temp;
                //outputpath_t = outputPath + "\\1.pdf";
                //temp = outputPath + "\\temp.pdf";
                //foreach (FileInfo f in l)
                //{
                //    if (l.IndexOf(f) == 0)
                //    {
                //        path1 = f.FullName;
                //        continue;
                //    }
                //    path2 = f.FullName;
                //    if (File.Exists(outputpath_t))
                //        File.Delete(outputpath_t);
                //    //comPDF(path1, path2, outputpath_t);
                //    FileInfo tempf = new FileInfo(outputpath_t);
                //    if (File.Exists(temp))
                //        File.Delete(temp);
                //    tempf.CopyTo(temp);
                //    path1 = temp;
                //}
                label3.Text = "正在清理临时文件";
                foreach (FileInfo f in l)
                    f.Delete();
                Directory.Delete(outputPath + "\\temp");
                progressBar1.Value++;
            }
            
            label3.Text = "已完成";
        }
    }
}