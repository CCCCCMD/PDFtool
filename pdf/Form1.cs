using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using SharpIO = PdfSharp.Pdf.IO;
using Sharp = PdfSharp.Pdf;
using Document = iTextSharp.text.Document;
using iTextSharp.text.pdf;

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

        private void comPDF0(String filename1, String filename2, String out3)
        {
            Sharp.PdfDocument inputDocument1 = SharpIO.PdfReader.Open(filename1, SharpIO.PdfDocumentOpenMode.Import);
            Sharp.PdfDocument inputDocument2 = SharpIO.PdfReader.Open(filename2, SharpIO.PdfDocumentOpenMode.Import);
            Sharp.PdfDocument outputDocument = new Sharp.PdfDocument();
            // Show consecutive pages facing. Requires Acrobat 5 or higher.
            outputDocument.PageLayout = Sharp.PdfPageLayout.TwoColumnLeft;
            int count = Math.Max(inputDocument1.PageCount, inputDocument2.PageCount);
            for (int idx = 0; idx < count; idx++)
            {
                Sharp.PdfPage page1 = inputDocument1.PageCount > idx ?
                  inputDocument1.Pages[idx] : new Sharp.PdfPage();
                Sharp.PdfPage page2 = inputDocument2.PageCount > idx ?
                  inputDocument2.Pages[idx] : new Sharp.PdfPage();

                // Add both pages to the output document
                page1 = outputDocument.AddPage(page1);
                page2 = outputDocument.AddPage(page2);


            }
            string filename = out3;
            outputDocument.Save(filename);
        }

        private void comPDF0(List<string> filelist, string path0)
        {
            Sharp.PdfDocument outputDocument = new Sharp.PdfDocument();
            outputDocument.PageLayout = Sharp.PdfPageLayout.TwoColumnLeft;
            foreach (string f in filelist)
            {
                Sharp.PdfDocument inputDocument = SharpIO.PdfReader.Open(f, SharpIO.PdfDocumentOpenMode.Import);
                int count = inputDocument.PageCount;
                for (int idx = 0; idx < count; idx++)
                {
                    Sharp.PdfPage page = inputDocument.Pages[idx];
                    outputDocument.AddPage(page);
                }
                inputDocument.Close();
            }
            string filename = path0;
            outputDocument.Save(filename);

        }

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
            dataGridView1.Rows.Clear();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            while (lst.Count == 0)
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    inputPath = dialog.SelectedPath;
                    textBox1.Text = inputPath;
                    getdir(inputPath, lst);
                    if (lst.Count == 0)
                    {
                        MessageBox.Show("该目录下没有文件，请重新选择！", "警告");
                    }
                    foreach (FileInfo f in lst)
                    {
                        String type = "其他文档";
                        if (f.Extension == ".docx" || f.Extension == ".doc")
                            type = "Word文档";
                        if (f.Extension == ".xlsx" || f.Extension == ".xls")
                            type = "Excel文档";
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

        void newth()
        {
            int flag = 1;
            if (radioButton2.Checked)
            {
                flag = 0;
                progressBar1.Maximum++;
            }
            foreach (FileInfo f in lst)
            {
                string ex = f.Extension.Replace(".", "");
                if (ex == "docx" || ex == "doc" || ex == "txt")
                    WordToPDF(f.FullName, toPDF(f, ex, flag));
                if (ex == "xlsx" || ex == "xls")
                    ExcelToPDF(f.FullName, toPDF(f, ex, flag));
                progressBar1.Value++;
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
                comPDF(pdf, outputPath + "\\" + dir.Name + ".pdf");
                foreach (FileInfo f in l)
                    f.Delete();
                Directory.Delete(outputPath + "\\temp");
                progressBar1.Value++;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            progressBar1.Value = progressBar1.Minimum;
            label3.Text = "开始转换";
            int flag = 1;
            if (radioButton2.Checked)
            {
                flag = 0;
                progressBar1.Maximum++;
            }
            foreach (FileInfo f in lst)
            {
                string ex = f.Extension.Replace(".", "");
                if (ex == "docx" || ex == "doc" || ex == "txt")
                    WordToPDF(f.FullName, toPDF(f, ex, flag));
                if (ex == "xlsx" || ex == "xls")
                    ExcelToPDF(f.FullName, toPDF(f, ex, flag));
                progressBar1.Value++;
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
                comPDF0(pdf, outputPath + "\\" + dir.Name + ".pdf");
                foreach (FileInfo f in l)
                    f.Delete();
                Directory.Delete(outputPath + "\\temp");
                progressBar1.Value++;
            }
            label3.Text = "已完成";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}