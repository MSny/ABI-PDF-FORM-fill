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
        string path = "C:\\Users\\Meir\\Downloads\\ABISample\\Outputsample.pdf";
        List<string[]> parsedCSVData = new List<string[]>();
        List<object[]> parsedPDFData = new List<object[]>();
        List<object> pdftextFields = new List<object>();
        dynamic pdfForm;
        dynamic JsonData;
        List<string> formsNtype = new List<string>();
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
                ReadExistingExcel();
            }

        }

        //Excel editing class
        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;
        public static void ReadExistingExcel()
        {
            string path = @"C:\Users\Meir\Downloads\ABISample\Copy of Boiler Batch OP-42 Form v01,2013 BLANK.XLSX";
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get the allready exists sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("OP-42");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;
            // Edit values
            mWSheet1.Cells[8, 7] = "Codeglomerate";
            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;
            // bottom data 
            for (int index = 1; index < 15; index++)
            {
                mWSheet1.Cells[34 + index, 2] = "Brklyn" + index;
                mWSheet1.Cells[34 + index, 3] = "Device" + index;
                mWSheet1.Cells[34 + index, 4] = "MD" + index;
                mWSheet1.Cells[34 + index, 5] = "Serial" + index;
                mWSheet1.Cells[34 + index, 6] = "House" + index;
                mWSheet1.Cells[34 + index, 7] = "Address" + index;
                mWSheet1.Cells[34 + index, 8] = "Block" + index;
                mWSheet1.Cells[34 + index, 9] = "Lot" + index;
                mWSheet1.Cells[34 + index, 10] = "Insp date" + index;
            }
            mWorkBook.SaveAs(@"C:\Users\Meir\Downloads\ABISample\newABISample", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing,
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

                String[] checkboxstates = fields.GetAppearanceStates("topmostSubform[0].Page1[0].CheckBox2[0]");

                // set form fields
                foreach (KeyValuePair<string, AcroFields.Item> kvp in fields.Fields)
                {
                    if (inputFileName == "bo9.pdf")
                    {
                        count++;
                        fields.SetField("topmostSubform[0].Page1[0].Boiler_Make___Model[0]", JsonData.person.name.ToString());
                        fields.SetField("topmostSubform[0].Page1[0].CheckBox2[0]", "1");
                        fields.SetField("topmostSubform[0].Page1[0].LocationFloor[0]", "1");
                    }
                    if (inputFileName == "bo13e.pdf")
                    {
                        count++;
                        fields.SetField("topmostSubform[0].Page1[0].Boiler_Make___Model[0]", JsonData.person.name.ToString());
                        fields.SetField("topmostSubform[0].Page1[0].CheckBox2[0]", "1");
                        fields.SetField("topmostSubform[0].Page1[0].LocationFloor[0]", "1");
                    }
                    if (inputFileName == "bo13.pdf")
                    {
                        count++;
                        fields.SetField("topmostSubform[0].Page1[0].Boiler_Make___Model[0]", JsonData.person.name.ToString());
                        fields.SetField("topmostSubform[0].Page1[0].CheckBox2[0]", "1");
                        fields.SetField("topmostSubform[0].Page1[0].LocationFloor[0]", "1");
                    }
                    if (inputFileName == "ABI FORM.pdf")
                    {
                        count++;
                        fields.SetField(formsNtype[count - 1], JsonData.person.name.ToString());
                        // ABI FORM CHECKBOX FILL VALUE 
                        fields.SetField("Check Box6", "Yes");
                        fields.SetField("Check Box7", "true");
                        

                        Console.Write(formsNtype[count - 1] + System.Environment.NewLine);
                    }
                }


              

                // flatten form fields and close document
                stamper.FormFlattening = true;
                stamper.Close();
            }      
        }
    }
}


