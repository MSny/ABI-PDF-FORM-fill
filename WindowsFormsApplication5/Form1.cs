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


using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.Windows.Forms;

using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;

namespace WindowsFormsApplication5
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        string PDFPath;
        string CSVPath;
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
            this.PDFPath = openFileDialog1.FileName;
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
                    count++;
                   // fields.SetField(formsNtype[count-1], parsedCSVData[count][0]);
                    fields.SetField("topmostSubform[0].Page1[0].Boiler_Make___Model[0]", JsonData.person.name.ToString());
                    // ABI FORM CHECKBOX FILL VALUE fields.SetField("Check Box6", "Yes");
                    fields.SetField("topmostSubform[0].Page1[0].CheckBox2[0]", "1");
                    fields.SetField("topmostSubform[0].Page1[0].LocationFloor[0]", "1");
                    Console.Write(formsNtype[count - 1] + System.Environment.NewLine);
                }

               
               // fields.SetField("BOILER", "HOT");
                //fields.SetField("", "12345");
                //fields.SetField("email", "johndoe@xxx.com");

                // flatten form fields and close document
                stamper.FormFlattening = true;
                stamper.Close();
            }
            /*
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();
            path = saveFileDialog1.FileName;
            // 3. Filling the acroform fields...

            foreach (Field currentField in pdfForm.Fields.Values)
                {
                    Console.Write(currentField);
                if (currentField is org.pdfclown.documents.interaction.forms.CheckBox)
                    {
                    org.pdfclown.documents.interaction.forms.CheckBox localCast = (org.pdfclown.documents.interaction.forms.CheckBox)currentField;
                    localCast.Checked = true;
                    continue;
                    }
                else if (currentField is org.pdfclown.documents.interaction.forms.RadioButton)
                             {
                    org.pdfclown.documents.interaction.forms.RadioButton localCast = (org.pdfclown.documents.interaction.forms.RadioButton)currentField;
                    currentField.Value = ((DualWidget)currentField.Widgets[0]).WidgetName;
                             } // Selects the first widget in the group.
                else if (currentField is ChoiceField)
                             {
                                 currentField.Value = ((ChoiceField)currentField).Items[0].Value;
                             } // Selects the first item in the list.
                             
                else
                             {
                                 CSVData = parsedCSVData[count][0];  // Arbitrary value (just to get something to fill with).
                                 Console.Write(CSVData);
                                 currentField.Value = CSVData;
                                 count++; 
                             }
                             
                    }
                
            
            */
        }

       
    }
    }


