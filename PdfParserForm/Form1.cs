using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Threading;

namespace PdfParserForm
{
    public partial class Form1 : Form
    {
        string input_file = "";
        string output_location = "";
        string output_location_processing = "";


        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                label4.Visible = true;
                label4.Text= openFileDialog1.FileName;
                input_file = openFileDialog1.FileName;

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                label5.Visible = true;
                label5.Text = folderBrowserDialog1.SelectedPath;
                output_location = folderBrowserDialog1.SelectedPath;

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                
                label6.Visible = true;
                label6.Text = folderBrowserDialog1.SelectedPath;
                output_location_processing = folderBrowserDialog2.SelectedPath;

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                label8.Visible = true;
                label8.Text = "Please Wait for 10-15 seconds...";
                Thread fileProcessingThread = new Thread(()=>DoFileProcessingTask(input_file, output_location, output_location_processing));
                fileProcessingThread.Start();             



            }
            catch(Exception exception)
            {
                label8.Text = exception.ToString();

            }
            
        }

        private void DoFileProcessingTask(string input_file, string output_location, string processing_location)
        {
            Pdfconverter pdfconverter = new Pdfconverter();
            pdfconverter.start(input_file, output_location, output_location_processing);
            label8.Text = "Excel Files Generated Successfully";
        }
    }
}
