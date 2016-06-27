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
        string[] files;
        string PDFPath;
        string CSVPath;
        Document PDFfields;
        string outputPath;
        List<string[]> parsedCSVData = new List<string[]>();
        List<string[]> parsedPDFData = new List<string[]>();
        string CSVData;
        int count = 5;
        public Form1()
        {
            InitializeComponent();
        }


        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //this.Activate();
            // files = openFileDialog1.FileNames; 
           // DialogResult result = openFileDialog1.ShowDialog();
            this.PDFPath = openFileDialog1.FileName;
            org.pdfclown.files.File file = new org.pdfclown.files.File(this.PDFPath);
            Document PDFfields = file.Document;
            // 2. Get the acroform!
            org.pdfclown.documents.interaction.forms.Form form = PDFfields.Form;
            Console.Write(PDFfields);
            if (!form.Exists())
                Console.WriteLine("No form available");
            else
            {

                // foreach (string field in document.Form.Fields.Values)
                //   parsedPDFData.Add();
                Console.WriteLine("Please select what data you would like");
                // 3. Filling the acroform fields...


                //foreach (KeyValuePair<string, Field> currentField in form.Fields)
                foreach (Field currentField in form.Fields.Values)
                {
                    Console.Write(currentField);
                    if (currentField is org.pdfclown.documents.interaction.forms.CheckBox)
                    {
                        org.pdfclown.documents.interaction.forms.CheckBox localCast = (org.pdfclown.documents.interaction.forms.CheckBox)currentField;
                        localCast.Checked = true;
                        continue;
                        //     String value;
                        /*     if (field is org.pdfclown.documents.interaction.forms.RadioButton)
                             {
                                 value = ((DualWidget)field.Widgets[0]).WidgetName;
                             } // Selects the first widget in the group.
                             else if (field is ChoiceField)
                             {
                                 value = ((ChoiceField)field).Items[0].Value;
                             } // Selects the first item in the list.
                             
                             else
                             {
                                // CSVData = parsedCSVData[count][0];  // Arbitrary value (just to get something to fill with).
                                // Console.Write(CSVData);
                                // field.Value = CSVData;
                                // count++;
                             }
                             */
                    }
                }
                    // 4. Serialize the PDF file!


                try
                {
                    //outputPath = @"C:\Users\Meir\Downloads\ABISample\OutputSample.pdf";
                    file.Save(SerializationModeEnum.Standard);
                }
                catch (Exception z)
                {
                    Console.WriteLine("File writing failed: " + z.Message);
                    Console.WriteLine(z.StackTrace);
                }
                Console.WriteLine("\nFile Created: ");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            openFileDialog2.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            //this.Activate();
            //files = openFileDialog1.FileNames;
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
          /*  if (!form.Exists())
                Console.WriteLine("No form available");
            else
            {

                // foreach (string field in document.Form.Fields.Values)
                //   parsedPDFData.Add();
                Console.WriteLine("Please select what data you would like");
                // 3. Filling the acroform fields...


          
                foreach (Field currentField in form.Fields.Values)
                {
                    Console.Write(currentField);
                    if (currentField is org.pdfclown.documents.interaction.forms.CheckBox)
                    {
                        org.pdfclown.documents.interaction.forms.CheckBox localCast = (org.pdfclown.documents.interaction.forms.CheckBox)currentField;
                        localCast.Checked = true;
                        continue;
                        //     String value;
                        /*     if (field is org.pdfclown.documents.interaction.forms.RadioButton)
                             {
                                 value = ((DualWidget)field.Widgets[0]).WidgetName;
                             } // Selects the first widget in the group.
                             else if (field is ChoiceField)
                             {
                                 value = ((ChoiceField)field).Items[0].Value;
                             } // Selects the first item in the list.
                             
                             else
                             {
                                // CSVData = parsedCSVData[count][0];  // Arbitrary value (just to get something to fill with).
                                // Console.Write(CSVData);
                                // field.Value = CSVData;
                                // count++;
                             }
                             
                    }
                }
                // 4. Serialize the PDF file!


                try
                {
                    //outputPath = @"C:\Users\Meir\Downloads\ABISample\OutputSample.pdf";
                    file.Save(SerializationModeEnum.Standard);
                }
                catch (Exception z)
                {
                    Console.WriteLine("File writing failed: " + z.Message);
                    Console.WriteLine(z.StackTrace);
                }
                Console.WriteLine("\nFile Created: ");
            }
        }*/
    }
}
    }

