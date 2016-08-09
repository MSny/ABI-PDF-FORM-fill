using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Configuration;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.Windows.Forms;


namespace WindowsFormsApplication5
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        string PDFPath;
        string CSVPath;
        string inputFileName;
        string extension;
        string currentItem;
        string currentCheckbox;
        string path = "C:\\Users\\Meir\\Downloads\\ABISample\\Outputsample.pdf";
        List<string[]> parsedCSVData = new List<string[]>();
        List<object[]> parsedPDFData = new List<object[]>();
        List<object> pdftextFields = new List<object>();
        dynamic pdfForm;
        dynamic JsonData;
        List<string> formsNtype = new List<string>();
        //List<string> checkboxCount = new List<string>();
        int count = 0;
        public Form1()
        {
            InitializeComponent();
        }

        //Select Pdf 
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            PdfReader.unethicalreading = true;
            inputFileName = openFileDialog1.SafeFileName;
            extension = Path.GetExtension(inputFileName);
            this.PDFPath = openFileDialog1.FileName;

            if (extension == ".pdf")
            {
                PdfReader reader = new PdfReader(this.PDFPath);
                AcroFields form = reader.AcroFields;
                pdfForm = reader;
                // 2. Get the acroform!

                if (pdfForm == null)
                    Console.WriteLine("No form available");
                else
                {
                    try
                    {
                        foreach (KeyValuePair<string, AcroFields.Item> kvp in form.Fields)
                        {
                            switch (form.GetFieldType(kvp.Key))
                            {
                                case AcroFields.FIELD_TYPE_CHECKBOX:
                                  //  string translatedCheckboxName = form.GetTranslatedFieldName(kvp.Key);
                                    //checkboxCount.Add(translatedCheckboxName);
                                  //  break;
                                case AcroFields.FIELD_TYPE_COMBO:
                                case AcroFields.FIELD_TYPE_LIST:
                                case AcroFields.FIELD_TYPE_RADIOBUTTON:
                                case AcroFields.FIELD_TYPE_NONE:
                                case AcroFields.FIELD_TYPE_PUSHBUTTON:
                                case AcroFields.FIELD_TYPE_SIGNATURE:
                                case AcroFields.FIELD_TYPE_TEXT:
                                    int fileType = form.GetFieldType(kvp.Key);
                                    string fieldValue = form.GetField(kvp.Key);
                                    string translatedFileName = form.GetTranslatedFieldName(kvp.Key);
                                    formsNtype.Add(translatedFileName);
                                    break;
                            }
                        }

                    }
                    catch
                    {
                    }
                    /*finally
                {
                reader.Close();
                }*/
                    Console.WriteLine("Form Selected");
                    // 3. Filling the acroform fields...
                }
            }
            else if (extension == ".XLSX")
            {
                ReadExistingExcel(JsonData);
            }

        }




        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }

        //Select Data
        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            this.CSVPath = openFileDialog2.FileName;
            JObject parsedData = JObject.Parse(File.ReadAllText(CSVPath));

            JsonData = parsedData;
            Console.WriteLine(parsedData);

        }

        //Excel editing class
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet2;
        private static Microsoft.Office.Interop.Excel.Application oXL;
        public static void ReadExistingExcel(dynamic data)
        {
            string path = @"C:\Users\Meir\Downloads\ABISample\Copy of Boiler Batch OP-42 Form v01,2013 BLANK.XLSX";
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get the existing sheets
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("OP-42");
            mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("csv");
            // mWSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("csv");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            // Edit values for OP-42
            mWSheet1.Cells[5, 20] = data.Boiler.id.ToString();
            mWSheet1.Cells[6, 21] = data.customerInfo.date.ToString();
            mWSheet1.Cells[8, 7] = data.company.ToString();
            mWSheet1.Cells[10, 7] = data.customerInfo.name.ToString();
            mWSheet1.Cells[12, 6] = data.Boiler.id.ToString();
            mWSheet1.Cells[12, 9] = data.customerInfo.job.ToString();
            mWSheet1.Cells[14, 7] = data.customerInfo.address.ToString();
            //mWSheet1.Cells[15, 20] = data.number.ToString();
            //mWSheet1.Cells[17, 21] = data.number.ToString();
            mWSheet1.Cells[17, 10] = data.customerInfo.date.ToString();
            mWSheet1.Cells[16, 5] = data.customerInfo.name_2.ToString();
            mWSheet1.Cells[18, 10] = data.customerInfo.phone.ToString();
            mWSheet1.Cells[18, 5] = data.customerInfo.email.ToString();
            mWSheet1.Cells[30, 10] = data.customerInfo.date.ToString();
            // Edit values for csv

            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;
            // bottom data 
            for (int index = 1; index < 10; index++)
            {
                mWSheet2.Cells[0 + index, 1] = data.csv.boro.ToString();
                mWSheet2.Cells[0 + index, 2] = data.csv.device.ToString();
                mWSheet2.Cells[0 + index, 3] = data.csv.md.ToString();
                mWSheet2.Cells[0 + index, 4] = data.csv.serial.ToString();
                mWSheet2.Cells[0 + index, 5] = data.csv.house.ToString();
                mWSheet2.Cells[0 + index, 6] = data.csv.street.ToString();
                mWSheet2.Cells[0 + index, 7] = data.csv.block.ToString();
                mWSheet2.Cells[0 + index, 8] = data.csv.lot.ToString();
                mWSheet2.Cells[0 + index, 9] = data.csv.date.ToString();
                mWSheet2.Cells[0 + index, 10] = data.csv.j.ToString();
                mWSheet2.Cells[0 + index, 11] = data.csv.k.ToString();
                mWSheet2.Cells[0 + index, 12] = data.csv.l.ToString();
                mWSheet2.Cells[0 + index, 13] = data.csv.m.ToString();
                mWSheet2.Cells[0 + index, 14] = data.csv.n.ToString();
                mWSheet2.Cells[0 + index, 15] = data.csv.o.ToString();
                mWSheet2.Cells[0 + index, 16] = data.csv.p.ToString();
                mWSheet2.Cells[0 + index, 17] = data.csv.q.ToString();
                mWSheet2.Cells[0 + index, 18] = data.csv.r.ToString();
                mWSheet2.Cells[0 + index, 19] = data.csv.location.ToString();
                mWSheet2.Cells[0 + index, 20] = data.csv.t.ToString();
            }
            SaveFileDialog saveFileDialog2 = new SaveFileDialog
            {
                Filter = "xlsx files|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true
            };
            saveFileDialog2.ShowDialog();
            string savePath = saveFileDialog2.FileName;
            mWorkBook.SaveAs(savePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            mWorkBook.Close(true, "newABIExcel", false);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = "PDF files|*.pdf",
                DefaultExt = "pdf",
                AddExtension = true
            };
            saveFileDialog1.ShowDialog();
            path = saveFileDialog1.FileName;
            Console.Write(JsonData);
            using (PdfStamper stamper = new PdfStamper(pdfForm, new FileStream(path, FileMode.Create)))
            {
                AcroFields fields = stamper.AcroFields;
                // fields.GenerateAppearances = true;
               // String[] checkboxstates = fields.GetAppearanceStates("topmostSubform[0].Page1[0].CheckBox2[0]");


                // set form fields
                foreach (KeyValuePair<string, AcroFields.Item> kvp in fields.Fields)
                {
                    
                    count++;
                    if (formsNtype[count - 1].StartsWith("topmostSubform"))
                    {
                        Console.Write(formsNtype[count - 1].Substring(27, formsNtype[count - 1].Length - 27) + System.Environment.NewLine);
                        currentCheckbox = null;
                        currentItem = null;
                    }
                    else
                    {
                        Console.Write(formsNtype[count - 1] + System.Environment.NewLine);
                        currentCheckbox = "0";
                        currentItem = "0";
                    }
                    
                    switch (fields.GetFieldType(kvp.Key))
                    {
                        case AcroFields.FIELD_TYPE_CHECKBOX:
                             currentCheckbox = string.IsNullOrEmpty(currentCheckbox) ? formsNtype[count - 1].Substring(27, formsNtype[count - 1].Length - 27): fields.GetTranslatedFieldName(kvp.Key);
                            //currentCheckbox = fields.GetTranslatedFieldName(kvp.Key);
                            fields.SetField(currentCheckbox, JsonData[inputFileName][currentCheckbox].ToString());
                            break;
                        case AcroFields.FIELD_TYPE_COMBO:
                        case AcroFields.FIELD_TYPE_LIST:
                        case AcroFields.FIELD_TYPE_RADIOBUTTON:
                            String[] radioStates = fields.GetAppearanceStates("topmostSubform[0].Page1[0].RadioButtonList[0]");
                            string currentRadio =  formsNtype[count - 1].Substring(27, formsNtype[count - 1].Length - 27) ;
                            fields.SetField(currentRadio, JsonData[inputFileName][currentRadio].ToString());
                            break;
                        case AcroFields.FIELD_TYPE_NONE:
                        case AcroFields.FIELD_TYPE_PUSHBUTTON:
                        case AcroFields.FIELD_TYPE_SIGNATURE:
                        case AcroFields.FIELD_TYPE_TEXT:
                            currentItem = string.IsNullOrEmpty(currentItem) ? formsNtype[count - 1].Substring(27, formsNtype[count - 1].Length - 27): fields.GetTranslatedFieldName(kvp.Key);

                           // currentItem = fields.GetTranslatedFieldName(kvp.Key);
                             fields.SetField(currentItem, JsonData[inputFileName][currentItem].ToString());
                            break;

                           
                                //fields.GenerateAppearances = true;

                                // Console.Write(formsNtype[count - 1] + System.Environment.NewLine);
                            }
                    }
                    // flatten form fields and close document
                    stamper.FormFlattening = true;
                    stamper.Close();
                }
            }
        }
    }



