using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

//Tim Fairfield 3/4/22
//use this when scanning a bunch of mixed up papers you want to split for filing
//get pdfsharp from nuget
//based on http://www.pdfsharp.net/wiki/HelloWorld-sample.ashx


namespace pdfsplitter
{

    public partial class pdfsplitter : Form
    {
        public string workingFile { get; set; }
        PdfDocument workingoutputDocument = new PdfDocument();
        public pdfsplitter()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int idx;
            // Get a fresh copy of the sample PDF file
            //const string filename = "timProfile.pdf";
            string filename = workingFile;


            //File.Copy(Path.Combine(@"C:\inout", filename),
            //Path.Combine(Directory.GetCurrentDirectory(), filename), true);

            // Open the file
            PdfDocument inputDocument = PdfReader.Open(filename, PdfDocumentOpenMode.Import);

            string name = Path.GetFileNameWithoutExtension(filename);
            string dirName = Path.GetDirectoryName(filename);
            Directory.SetCurrentDirectory(dirName);

            bool isinSplitlist = true;//true for 0 and numbers in textbox otherwise not
            bool fillingfile = false;

            string split = textBox2.Text;
            List<string> list = new List<string>();
            list = split.Split(',').ToList();
            //foreach (var l in list)
            //{
            //    Console.WriteLine(l);
            //}

            ///make multiple lists which describe which pages are in each doc
            ///calculate pages in the list or allow user to select which specific pages  like 1-3,2-4,6-8 or a single number
            ///then open the doc use parameters for start and end pages and filenames and count through
            /// fmethod   (sourceDocname, list of pages to include, outDocname)



            //for (int idx = 0; idx < inputDocument.PageCount; idx++)
            //{

            //    if (list.Contains(idx.ToString())) //this is for single pages
            //        Console.WriteLine("contains "+ idx.ToString());

            //    {
            //        fillingfile = true;
            // writeSingleDoc(inputDocument, name);
            //    }
            //    if (idx>0 && list[idx-1].Contains("-")) //if there is a hyphen then it wont show above we need to find the first and last number and iterate
            //        {
            //        Console.WriteLine("contains " + list[idx-1]);
            //    }

            foreach (string entry in list)
            {

                if (!entry.Contains("-"))
                {
                    Console.WriteLine("contains " + entry);
                    //write_singlepage()
                    idx = int.Parse(entry)-1;
                    if (idx > inputDocument.PageCount || idx < 0)
                    {
                        MessageBox.Show("chosen page range not valid (check your page range selection) ");
                    }
                    else
                    {

                        writeSingleDoc(inputDocument, name, idx);
                    }
                }


                if (entry.Contains("-"))
                {

                    Console.WriteLine("contains " + entry);
                   string[] minmax = entry.Split('-');

                    try
                    {
                        int min = int.Parse(minmax[0]) - 1;
                        int max = int.Parse(minmax[1]) - 1;

                        if (max > inputDocument.PageCount)
                            max = inputDocument.PageCount;

                        if (min > inputDocument.PageCount || min < 0)
                        {
                            MessageBox.Show("chosen page range not valid (check your page range selection) ");

                        }
                        else
                        {
                            //get first and last number
                            //make a loop to write page starting and ending with last number 
                            writeMultipageDoc(inputDocument, name, min, max);
                        }

                    }

                    catch
                    {
                        MessageBox.Show("chosen page range not valid (check your page range selection) ");
                    }
                }


            }






        }

        private void writeSingleDoc(PdfDocument inputDocument, string name, int idx)
        {
            // Create new document
            PdfDocument outputDocument = new PdfDocument();
            outputDocument.Version = inputDocument.Version;
            outputDocument.Info.Title =
              String.Format("Page {0} of {1}", idx + 1, inputDocument.Info.Title);
            outputDocument.Info.Creator = inputDocument.Info.Creator;
            workingoutputDocument = outputDocument;
            // Add split page
            outputDocument.AddPage(inputDocument.Pages[idx]);
            //save that doc
            workingoutputDocument.Save(String.Format("{0} - Page {1}_split.pdf", name, idx + 1));
        }
        private void writeMultipageDoc(PdfDocument inputDocument, string name, int min, int max)
        {

            // Create new document
            PdfDocument outputDocument = new PdfDocument();
            outputDocument.Version = inputDocument.Version;
            outputDocument.Info.Title =
              String.Format("Page {0} of {1}", min + 1, inputDocument.Info.Title);
            outputDocument.Info.Creator = inputDocument.Info.Creator;
            workingoutputDocument = outputDocument;

            int idx = 0;
            for (idx = min; idx < max; idx++)
            { 
            // Add split page
            outputDocument.AddPage(inputDocument.Pages[idx]);
              }
            //save that doc
            workingoutputDocument.Save(String.Format("{0} - Page {1}_split.pdf", name, min+1));//name min+1 to correct of page 1 being actually page zero
        }



        private void button2_Click(object sender, EventArgs e)
        {
            //var fileContent = string.Empty;
            //var filePath = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    workingFile = openFileDialog.FileName;
                    textBox1.Text = workingFile;
                    textBox1.Refresh();



                    ////Read the contents of the file into a stream
                    //var fileStream = openFileDialog.OpenFile();


                    ////using (StreamReader reader = new StreamReader(fileStream))
                    ////{
                    ////    fileContent = reader.ReadToEnd();
                    //}
                }
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
