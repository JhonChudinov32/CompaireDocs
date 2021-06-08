using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using OfficeWord = Microsoft.Office.Interop.Word;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Color = System.Drawing.Color;

namespace CompaireDocs
{
    public partial class CompairDoc : Form
    {
        private string DS;
        private string files;
        public CompairDoc()
        {
            InitializeComponent();
            comboBox1.Items.AddRange(new string[] {"2020", "2021"});
            comboBox1.SelectedIndexChanged += ComboBox1_SelectedIndexChanged;
        }
        private void DGPaint()
        {
            for (int i = 0; i <= dataGridView1.ColumnCount - 1; i++)
            {
                for (int j = 0; j <= dataGridView1.RowCount - 1; j++)
                {
                    dataGridView1[i, j].Style.BackColor = Color.White;
                    dataGridView1[i, j].Style.ForeColor = Color.Black;
                }
            }
            for (int i = 0; i <= dataGridView1.ColumnCount - 1; i++)
            {
                for (int j = 0; j <= dataGridView1.RowCount - 2; j++)
                {
                    if (dataGridView1[i, j].Value.ToString().IndexOf("<del>")!= -1)
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                       // dataGridView1[i, j].Style.BackColor = Color.AliceBlue;
                      //  dataGridView1[i, j].Style.ForeColor = Color.Red;
                    }
                    if (dataGridView1[i, j].Value.ToString().IndexOf("<ins>")!= -1)
                    {
                        dataGridView1[i, j].Style.BackColor = Color.AliceBlue;
                        dataGridView1[i, j].Style.ForeColor = Color.Red;
                    }
                }
            }
            
        }
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedState = comboBox1.SelectedItem.ToString();
        }
        private void Parser(string path)
        {
            object FileName = path;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            OfficeWord.Application app = new OfficeWord.Application();
            OfficeWord.Document doc = null;
            OfficeWord.Range range = null;
            try
            {
                doc = app.Documents.Open(ref FileName, ref MissingObj, ref rOnly, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj);

                object StartPosition = 0;
                object EndPositiojn = doc.Characters.Count;
                range = doc.Range(ref StartPosition, ref EndPositiojn);

                // Получение основного текста со страниц (без учёта сносок и колонтитулов)
                string MainText = (range == null || range.Text == null) ? null : range.Text;
                if (MainText != null)
                {

                    /* Обработка основного текста документа*/
                    textBox1.Text = MainText;
                }
                // Получение текста из нижних и верхних колонтитулов
                foreach (OfficeWord.Section section in doc.Sections)
                {
                    // Нижние колонтитулы
                    foreach (OfficeWord.HeaderFooter footer in section.Footers)
                    {
                        string FooterText = (footer.Range == null || footer.Range.Text == null) ? null : footer.Range.Text;
                        if (FooterText != null)
                        {
                            /* Обработка текста */
                            //textBox1.Text += FooterText;
                        }

                    }
                    // Верхние колонтитулы
                    foreach (OfficeWord.HeaderFooter header in section.Headers)
                    {
                        string HeaderText = (header.Range == null || header.Range.Text == null) ? null : header.Range.Text;
                        if (HeaderText != null)
                        {
                            /* Обработка текста */
                        }
                    }


                }
                // Получение текста сносок
                if (doc.Footnotes.Count != 0)
                {
                    foreach (OfficeWord.Footnote footnote in doc.Footnotes)
                    {
                        string FooteNoteText = (footnote.Range == null || footnote.Range.Text == null) ? null : footnote.Range.Text;
                        if (FooteNoteText != null)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                /* Обработка исключений */
                Console.WriteLine(ex.Message);
            }
            finally
            {
                /* Очистка неуправляемых ресурсов */
                if (doc != null)
                {
                    doc.Close(ref SaveChanges);
                }
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }
        private void Parser1(string path)
        {
            object FileName = path;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            OfficeWord.Application app = new OfficeWord.Application();
            OfficeWord.Document doc = null;
            OfficeWord.Range range = null;
            try
            {
                doc = app.Documents.Open(ref FileName, ref MissingObj, ref rOnly, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj);

                object StartPosition = 0;
                object EndPositiojn = doc.Characters.Count;
                range = doc.Range(ref StartPosition, ref EndPositiojn);

                // Получение основного текста со страниц (без учёта сносок и колонтитулов)
                string MainText = (range == null || range.Text == null) ? null : range.Text;
                if (MainText != null)
                {

                    /* Обработка основного текста документа*/
                    textBox2.Text = MainText;
                }
                // Получение текста из нижних и верхних колонтитулов
                foreach (OfficeWord.Section section in doc.Sections)
                {
                    // Нижние колонтитулы
                    foreach (OfficeWord.HeaderFooter footer in section.Footers)
                    {
                        string FooterText = (footer.Range == null || footer.Range.Text == null) ? null : footer.Range.Text;
                        if (FooterText != null)
                        {
                            /* Обработка текста */
                            //textBox1.Text += FooterText;
                        }

                    }
                    // Верхние колонтитулы
                    foreach (OfficeWord.HeaderFooter header in section.Headers)
                    {
                        string HeaderText = (header.Range == null || header.Range.Text == null) ? null : header.Range.Text;
                        if (HeaderText != null)
                        {
                            /* Обработка текста */
                        }
                    }


                }
                // Получение текста сносок
                if (doc.Footnotes.Count != 0)
                {
                    foreach (OfficeWord.Footnote footnote in doc.Footnotes)
                    {
                        string FooteNoteText = (footnote.Range == null || footnote.Range.Text == null) ? null : footnote.Range.Text;
                        if (FooteNoteText != null)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                /* Обработка исключений */
                Console.WriteLine(ex.Message);
            }
            finally
            {
                /* Очистка неуправляемых ресурсов */
                if (doc != null)
                {
                    doc.Close(ref SaveChanges);
                }
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }
        public void AddGridParam(string[] N, DataGridView Grid)
        {
            
            //пока столбцов не будет достаточное количество добавляем их
            while (N.Length > Grid.ColumnCount)
            {
                //если колонок нехватает добавляем их пока их будет хватать
                Grid.Columns.Add("", "");
            }

            //заполняем строку
            Grid.Rows.Add(N);
        }
        private bool CompareFile(string Path2, string Path1)
        {
         
                int file1byte;
                int file2byte;
          

                FileStream fs1 = new FileStream(Path1, FileMode.Open);
                FileStream fs2 = new FileStream(Path2, FileMode.Open);

                do
                {
                    file1byte = fs1.ReadByte();
                    file2byte = fs2.ReadByte();
                }
                while ((file1byte == file2byte) && (file1byte != -1));

                fs1.Close();
                fs2.Close();

                return ((file1byte - file2byte) == 0);
            
          
    
        }
        private string CompareFileWord1(string Path2, string Path1, string filefolder)
        {
            //create Word application
            var app = new OfficeWord.Application
            {
                DisplayAlerts = OfficeWord.WdAlertLevel.wdAlertsNone
            };
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object AddToRecent = false;
            object Visible = true;

            try
            {
              
                    //try open signed file 
                    OfficeWord.Document docZero = app.Documents.Open(Path2, ref missing, ref readOnly, ref AddToRecent, Visible: ref Visible);

                    docZero.Final = false;
                    docZero.TrackRevisions = true;
                    docZero.ShowRevisions = true;
                    docZero.PrintRevisions = true;

                    //compare file from card and signed file
                    docZero.Compare(Path1, missing, OfficeWord.WdCompareTarget.wdCompareTargetCurrent, true, false, false, false, false);

                    string fileName = textBox3.Text + "/";

                    //save file of compare
                    docZero.SaveAs2(fileName);
                    docZero.Close();

                    app.Quit();

                    return fileName;
              
            }

            catch
            {
                app.Quit();
                return "";
            }


        }
        private string CompareFileWord(string Path2, string Path1, string filefolder)
        {
            //create Word application
            var app = new OfficeWord.Application
            {
                DisplayAlerts = OfficeWord.WdAlertLevel.wdAlertsNone
            };
            object missing = System.Reflection.Missing.Value;
                object readOnly = false;
                object AddToRecent = false;
                object Visible = true;

                try
                {
               
                    //try open signed file 
                    OfficeWord.Document docZero = app.Documents.Open(Path2, ref missing, ref readOnly, ref AddToRecent, Visible: ref Visible);

                    docZero.Final = false;
                    docZero.TrackRevisions = true;
                    docZero.ShowRevisions = true;
                    docZero.PrintRevisions = true;

                    //compare file from card and signed file
                    docZero.Compare(Path1, missing, OfficeWord.WdCompareTarget.wdCompareTargetCurrent, true, false, false, false, false);
                    string name = "Результат" + "-" + System.IO.Path.GetFileName(Path1);
                    string fileName = textBox3.Text + "/" + name;

                    //save file of compare
                    docZero.SaveAs2(fileName);
                    docZero.Close();

                    Take_Compare(fileName);
                    app.Quit();

                    return fileName;
              
                }

                catch
                {
                    app.Quit();
                    return "";
                }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            //Открываем диалоговое окно
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //поиск с директории диска С
                openFileDialog.InitialDirectory = "C:/";
                //включаем возможность выбора несколько файлов
                openFileDialog.Multiselect = true;
                //включаем фильтр только word
                openFileDialog.Filter = "word files (*.docx)|*.docx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var Filenames = openFileDialog.FileNames;

                    //цикл открытия файлов поочередно
                    foreach (var fil in Filenames)
                    {
                        //цикл открытия файлов поочередно
                        Parser1(fil);
                        files = fil;
                    }
                }
            }
        }
        private void Button6_Click(object sender, EventArgs e)
        {
           dataGridView1.Rows.Clear();
           CompareFile(DS, files);
           CompareFileWord(DS, files, textBox3.Text + "/");
           ExRezult();
            DGPaint();
        }
        private string Take_Compare(string filePath)
        {
            string resultText = "";
            using (WordprocessingDocument wordprocDoc = WordprocessingDocument.Open(filePath, true))
            {
                Body body = wordprocDoc.MainDocumentPart.Document.Body;
                StringBuilder Result = new StringBuilder();
                //take each paragraph which contain text
                IEnumerable<Paragraph> paragraphs = body.Elements<Paragraph>().Where(paragrahp => paragrahp.InnerText != "");
                List<Paragraph> paragraphsList = paragraphs.ToList();
                string text = "";
                string text2 = "";
                string res = "";
                //Здесь вызов методов
                res = res + Result.ToString();

                //Здесь вызов методов
                int delFlag = 0;
                for (var i = 0; i < paragraphsList.Count(); i++)
                {
                   
                    //take paragraph which have local name "del"(this string contain change text)
                    if (paragraphsList[i].ChildElements.Where(child => child.LocalName == "del").Count() != 0 )
                    {
                        // if paragraph before not "del" paragraph, add as context 
                        if (delFlag == 0)
                        {
                            if (i > 0)
                            {
                                text = paragraphsList[i - 1].InnerText;
                                text2 = paragraphsList[i - 1].InnerText;
                            }
                            delFlag = 1;
                        }
                        //take text from "del" paragraph
                        foreach (OpenXmlElement child in paragraphsList[i].ChildElements)
                        {

                            if (child.LocalName == "del")
                            {

                                text = text + " " + " <del> " + child.InnerText + " </del> ";
                            }
                            else if (child.LocalName == "ins")
                            {
                                text2 = text2 + " " + " <ins> " + child.InnerText + " </ins> ";
                            }
                        }
                    }
                    else
                    {
                        //if before usual paragraph was "del" paragraph, add this paragraph and finish process
                        if (delFlag == 1)
                        {
                            text =  text + " " + paragraphsList[i].InnerText;
                            text2 = text2 + " " + paragraphsList[i].InnerText;
                            res = "";
                            res = res + Result.ToString();
                            delFlag = 0;
                            // заполняем datagrid 
                            string[] strokaEshe = { text, text2 };
                            AddGridParam(strokaEshe, dataGridView1);
                        }
                    }
                }
                if (delFlag == 1)
                {
                    //if file end on usual paragraph, after paragraph with "del"
                    text = (text + paragraphsList[paragraphsList.Count() - 1].InnerText).Replace(";", ",");
                    text2 = (text2 + paragraphsList[paragraphsList.Count() - 1].InnerText).Replace(";", ",");
                    res = "";
                    res= res + Result.ToString();
                    delFlag = 0;
        
                    // заполняем datagrid 
                    string[] strokaEshe = { text, text2 };
                    AddGridParam(strokaEshe, dataGridView1);
                }
                wordprocDoc.Close();
                //delete carriage return symbol
                if (Result.ToString().Length > 2)
                {
                    resultText = Result.ToString().Substring(0, Result.ToString().Length - 2);
                }
                else
                {
                    resultText = Result.ToString();
                }
                return resultText;
            }
        }
        private void Button7_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex >= 0 && checkBox1.Checked == false)
            {
                if (radioButton1.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/Д-01.docx";
                }
                else if (radioButton1.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/Д-01.docx";
                }
                else if (radioButton8.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-07.docx";
                }
                else if (radioButton8.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-07.docx";
                }
                else if (radioButton2.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-08.docx";
                }
                else if (radioButton2.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-08.docx";
                }
                else if (radioButton3.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-15.docx";
                }
                else if (radioButton3.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-15.docx";
                }
                else if (radioButton4.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-18.docx";
                }
                else if (radioButton4.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-18.docx";
                }
                else if (radioButton5.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-20.docx";
                }
                else if (radioButton5.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-20.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-22.docx";
                }
                else if (radioButton7.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-25.docx";
                }
                else if (radioButton7.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-25.docx";
                }

                else if (radioButton12.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-17.docx";
                }
                else if (radioButton12.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-17.docx";
                }
                else if (radioButton14.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020/С-24.docx";
                }
                else if (radioButton14.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021/С-24.docx";
                }
            }
            else if (comboBox1.SelectedIndex >= 0 && checkBox1.Checked == true)
            {
                if (radioButton1.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/Д-01.docx";
                }
                else if (radioButton1.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/Д-01.docx";
                }
                else if (radioButton2.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-08.docx";
                }
                else if (radioButton2.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-08.docx";
                }
                else if (radioButton3.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-15.docx";
                }
                else if (radioButton3.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-15.docx";
                }
                else if (radioButton4.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-18.docx";
                }
                else if (radioButton4.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-18.docx";
                }
                else if (radioButton5.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-20.docx";
                }
                else if (radioButton5.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-20.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-22.docx";
                }
                else if (radioButton6.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-22.docx";
                }
                else if (radioButton7.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-25.docx";
                }
                else if (radioButton7.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-25.docx";
                }
                else if (radioButton8.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-07.docx";
                }
                else if (radioButton8.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-07.docx";
                }
                else if (radioButton12.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-17.docx";
                }
                else if (radioButton12.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-17.docx";
                }
                else if (radioButton14.Checked == true && comboBox1.Text == "2020")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2020OB/С-24.docx";
                }
                else if (radioButton14.Checked == true && comboBox1.Text == "2021")
                {
                    DS = AppDomain.CurrentDomain.BaseDirectory + "dogovor2021OB/С-24.docx";
                }
            }
            Parser(DS);
            dataGridView1.Rows.Clear();
            
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfdoc = new SaveFileDialog
            {
                Filter = "Word Documents (*.doc)|*.docx",
                FileName = "Сравнение" + "-" + System.IO.Path.GetFileName(files) + "-" + DateTime.Today.Date.ToString("d")
            };
            if (sfdoc.ShowDialog() == DialogResult.OK)
            {
                CompareFileWord1(DS, files, sfdoc.FileName);// Here dataGridview1 is your grid view name
                Process.Start(sfdoc.FileName);
            }

        }
        private void ExRezult()
        {
            try
            {
                string fil = AppDomain.CurrentDomain.BaseDirectory + "Результат.xlsm";
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true
                };

                int StartCol = 1;
                int StartRow = 1;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Open(fil);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                int j = 0, i = 0;
                excel.Columns.ColumnWidth = 70;
                excel.Columns.WrapText = true;
                //Сохранение файла результат
                excel.Application.ActiveWorkbook.SaveAs(textBox3.Text + "/" + "Результат" + "-" + System.IO.Path.GetFileNameWithoutExtension(files) + ".xlsm");
                // excel.Application.ActiveWorkbook.SaveAs(textBox3.Text + "/" + "Результат" + "-" + System.IO.Path.GetFileNameWithoutExtension(files) + "-" + DateTime.Today.Date.ToString("d") + ".xlsm");
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    //myRange.Range["A1","B1000"].ClearContents();
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                    myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                    myRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                   
                    myRange.Font.Bold = true;
                    myRange.Font.Color = Color.Black;
                    myRange.Font.Size = 14;
                }
                StartRow++;
                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value ?? "";
                            myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            myRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                            myRange.Font.Bold = false;
                            myRange.Font.Size = 12;
                        }

                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.ToString());
                        }
                    }

                }
                //for (int l = 0; l < 100; l++ )
                //{
                //    TextCol((Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + l, StartCol + j]);
                //}
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void Button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                 textBox3.Text = FBD.SelectedPath;
               // MessageBox.Show("Папка выбрана");
            }
        }
    
    }
}
