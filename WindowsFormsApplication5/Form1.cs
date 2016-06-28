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


using org.pdfclown.documents;
using org.pdfclown.documents.contents.composition;
using org.pdfclown.documents.contents.entities;
using org.pdfclown.documents.contents.fonts;
using org.pdfclown.documents.contents.xObjects;
using org.pdfclown.documents.interaction.annotations;
using org.pdfclown.documents.interaction.forms;
using org.pdfclown.files;
using System.Windows.Forms;

using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;

namespace WindowsFormsApplication5
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        string PDFPath;
        string CSVPath;
        List<string[]> parsedCSVData = new List<string[]>();
        List<object[]> parsedPDFData = new List<object[]>();
        List<object> pdftextFields = new List<object>();
        dynamic pdfForm;
        string CSVData;
        int count = 5;
        public Form1()
        {
            InitializeComponent();
        }

        //Select Pdf 
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //this.Activate();
            // files = openFileDialog1.FileNames; 
           // DialogResult result = openFileDialog1.ShowDialog();
            this.PDFPath = openFileDialog1.FileName;
            org.pdfclown.files.File outputfile = new org.pdfclown.files.File(this.PDFPath);
            Document PDFfields = outputfile.Document;
            // 2. Get the acroform!
            org.pdfclown.documents.interaction.forms.Form form = PDFfields.Form;
            pdfForm = PDFfields.Form;
            Console.Write(PDFfields);
            if (!form.Exists())
                Console.WriteLine("No form available");
            else
            {
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
            TextFieldParser parser = new TextFieldParser(this.CSVPath);
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            while (!parser.EndOfData)
            {
                //Processing row
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    //TODO: Process field
                    parsedCSVData.Add(fields);
                }
            }

            parser.Close();
            Console.WriteLine(parsedCSVData);

        }



        private void button3_Click(object sender, EventArgs e)
        {
          
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
                               /*  CSVData = parsedCSVData[count][0];  // Arbitrary value (just to get something to fill with).
                                 Console.Write(CSVData);
                                 currentField.Value = CSVData;
                                 count++; */
                             }
                             
                    }
                
                // 4. Serialize the PDF file!


                try
                {
                    //outputPath = @"C:\Users\Meir\Downloads\ABISample\OutputSample.pdf";
                    pdfForm.File.Save(SerializationModeEnum.Standard);
                }
                catch (Exception z)
                {
                    Console.WriteLine("File writing failed: " + z.Message);
                    Console.WriteLine(z.StackTrace);
                }
                Console.WriteLine("\nFile Created: ");
            }
        }
    }

