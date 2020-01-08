using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop;
using System.Threading.Tasks;
using System.Diagnostics;

namespace PdfParserForm
{
    class Pdfconverter
    {
        String allLines = "";
        String[] allWords = { };
        ArrayList student_list = new ArrayList();
        List<string> myCollection = new List<string>();
        List<string> compressed = new List<string>();
        List<string> compressed1 = new List<string>();
        List<string> compressed2 = new List<string>();
        List<string> compressed3 = new List<string>();
        List<string> compressed4 = new List<string>();
        List<string> compressed5 = new List<string>();
        List<string> compressed6 = new List<string>();
        List<string> compressed7 = new List<string>();
        ArrayList studentlist = new ArrayList();
        string pattern = @"^(\d)(\d)[a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z](\d)(\d)(\d)$";
        string pattern2 = @"^(\d)(\d)[a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z][a-zA-Z](\d)(\d)(\d)";
        //string pattern3 = @"^(IT)[0-9]{1,3}$";
        string pattern3 = @"^(IT|MF)[0-9]{1,3}$";

        string pattern4 = @"^[A-Z]{2}[0-9]{3}$";
        //string pattern5 = @"^[A-Z]{2}$";
        //string pattern5 = @"^[AA|AB|BB|BC|CC|FF]";
        string pattern5 = @"^(AA|AB|BB|BC|CC|FF|CD|DD)$";
        string pattern6 = @"(IT|MF)[0-9]{1,3}";
        string number_pattern = @"^[0-9]";
        string six_digit_pattern = @"^[0-9]{6}$"; //Extreme Case 



        private List<String> removing_blank_words(String[] allWords, String processing_location)
        {
            string temp = "";
            List<string> myCollection = new List<string>();
            for (int i = 0; i < allWords.Length; i++)
            {
                temp = allWords[i].ToString();
                if (String.IsNullOrEmpty(temp))
                {

                }
                else
                {
                    myCollection.Add(temp);
                }

            }
            string path = processing_location+@"\Parse1_removing_blank_words.txt";
            File.WriteAllLines(path, myCollection, Encoding.UTF8);

            return myCollection;

        }

        private List<String> removing_waste_words(List<String> myCollection, String processing_location)
        {
            string temp = "";
            List<string> compressed = new List<string>();
            for (int i = 0; i < myCollection.Count; i++)
            {
                if (myCollection[i] == "Bachelor" || myCollection[i] == "Master")
                {
                    i = i + 21;
                    compressed.Add(myCollection[i]);
                }
                else if (myCollection[i] == "Jan" || myCollection[i] == "Dec" || myCollection[i] == "Feb" || myCollection[i] == "Apr" || myCollection[i] == "Mar" || myCollection[i] == "May" || myCollection[i] == "June" || myCollection[i] == "Jul" || myCollection[i] == "Aug" || myCollection[i] == "Sep" || myCollection[i] == "Oct" || myCollection[i] == "Nov")
                {
                    i = i + 25;
                    compressed.Add(myCollection[i]);
                }
                else if (myCollection[i] == "*")
                {

                }
                else if (myCollection[i] == ":")
                {

                }
                else
                {
                    compressed.Add(myCollection[i]);
                }
            }
            string path1 = processing_location+@"\Parse2_Compressed_removing_waste_words.txt";
            File.WriteAllLines(path1, compressed, Encoding.UTF8);

            return compressed;
        }

        private List<String> merging_name_roll(List<String> compressed, String processing_location)
        {
            string temp = "";
            List<string> compressed2 = new List<string>();
            for (int i = 0; i < compressed.Count; i++)
            {
                Match match = Regex.Match(compressed[i], pattern);
                if (match.Success)
                {
                    temp = compressed[i] + " " + compressed[i + 1] + " " + compressed[i + 2] + " " + compressed[i + 3];
                    Match match1 = Regex.Match(compressed[i + 4], pattern3);
                    if (!match1.Success)
                    {
                        compressed2.Add(temp);
                        i = i + 4;
                        compressed2.Add(compressed[i]);
                        temp = "";

                    }
                    else
                    {
                        temp = temp + " " + compressed[i + 4];
                        compressed2.Add(temp);
                        i = i + 5;
                        compressed2.Add(compressed[i]);
                        temp = "";

                    }

                }
                else
                {
                    compressed2.Add(compressed[i]);
                }

            }
            string path3 = processing_location+@"\Parse3_Name_Together.txt";
            File.WriteAllLines(path3, compressed2, Encoding.UTF8);

            return compressed2;
        }

        private List<String> merging_subject_name(List<String> compressed2, String processing_location)
        {
            string temp = "";
            List<string> compressed4 = new List<string>();
            Match match3;
            for (int i = 0; i < compressed2.Count; i++)
            {
                Match match2 = Regex.Match(compressed2[i], pattern4);
                if (match2.Success)
                {
                    compressed4.Add(compressed2[i]);
                    i = i + 1;
                    do
                    {
                        temp = temp + compressed2[i] + " ";
                        match3 = Regex.Match(compressed2[i + 1], pattern5);
                        i = i + 1;

                    } while (!match3.Success);
                    i = i - 1;
                    compressed4.Add(temp);
                    temp = "";
                }
                else
                {
                    compressed4.Add(compressed2[i]);

                }

            }
           
            

            string path5 = processing_location+@"\Parse4_All_Subject_Name_Together.txt";
            File.WriteAllLines(path5, compressed4, Encoding.UTF8);
            return compressed4;
        }
        private List<String> adding_dash(List<String> compressed4, String processing_location)
        {
          
            List<string> compressed5 = new List<string>();
            for (int i = 0; i < compressed4.Count; i++)
            {
                if (compressed4[i] == "---")
                {
                    compressed5.Add(compressed4[i]);
                    compressed5.Add("---");
                    compressed5.Add("---");

                }
                else
                {
                    compressed5.Add(compressed4[i]);
                }
            }
            string path6 = processing_location+@"\Parse5_dashadded.txt";
            File.WriteAllLines(path6, compressed5, Encoding.UTF8);

            return compressed5;
        }
        private List<String> next_row_name_revision(List<String> compressed5, String processing_location)
        {
            Match match_temp;
            string temp = "";
            List<string> compressed6 = new List<string>();
            for (int i = 0; i < compressed5.Count; i++)
            {
                Match match = Regex.Match(compressed5[i], pattern2);
                if (match.Success)
                {

                    Match match1 = Regex.Match(compressed5[i + 1], pattern4);
                    if (!match1.Success)
                    {
                        if((match_temp= Regex.Match(compressed5[i + 2], pattern4)).Success)
                        {
                            temp = temp + compressed5[i] + " " + compressed5[i + 1];
                            compressed6.Add(temp);
                            i = i + 1;
                            temp = "";
                        }
                        


                    }
                    else
                    {
                        compressed6.Add(compressed5[i]);

                    }
                }
                else
                {
                    compressed6.Add(compressed5[i]);
                }
            }
            string path7 = processing_location+@"\Parse6__name_Revision_next_line.txt";
            File.WriteAllLines(path7, compressed6, Encoding.UTF8);

            return compressed6;
        }
        private List<String> delete_duplicate_name(List<String> compressed6, String processing_location)
        {
            string temp = "";
            List<string> compressed7 = new List<string>();
            for (int i = 0; i < compressed6.Count; i++)
            {
                Match match = Regex.Match(compressed6[i], pattern2);
                if (match.Success)
                {
                    if (compressed6[i] == temp)
                    {
                        compressed7.Add(compressed6[i + 1]);
                        i = i + 1;
                        temp = " ";
                    }
                    else
                    {
                        temp = compressed6[i];
                        compressed7.Add(temp);
                    }


                }
                else
                {
                    compressed7.Add(compressed6[i]);
                }



            }
            compressed7.Add("ENDING");
            string path8 = processing_location+@"\Parse7_duplicate_name_delete_Revision.txt";
            File.WriteAllLines(path8, compressed7, Encoding.UTF8);
            
            return compressed7;
        }
        private void allstudentdetails(List<String> compressed2, String processing_location)
        {
            List<string> compressed3 = new List<string>();
            for (int i = 0; i < compressed2.Count; i++)
            {
                Match match2 = Regex.Match(compressed2[i], pattern2);
                if (match2.Success)
                {

                    compressed3.Add(compressed2[i]);
                }

            }

            string path4 = processing_location+@"\Parse8_Student_Details.txt";
            File.WriteAllLines(path4, compressed3, Encoding.UTF8);


        }


        private ArrayList make_student_objects_from_stringlist(List<String> compressed7)
        {

            int number_of_subjects = 0;
            int k = 0;
            Match subject;
            Match student_id;

            //do
            //{
            //    subject = Regex.Match(compressed7[k + 1], pattern4);
            //    if (subject.Success)
            //    {
            //        number_of_subjects = number_of_subjects + 1;
            //    }
            //    k++;
            //}
            //while (!(student_id = Regex.Match(compressed7[k + 1], pattern2)).Success);
            //String path5 = @"C:\Users\admin\Desktop\testing\Console_log.txt";
            //File.WriteAllText(path5, number_of_subjects.ToString(), Encoding.UTF8);

            int pdf_type=0;
            
            for (int z=0;z<compressed7.Count;z++)
            {
                if(compressed7[z]=="POINTS")
                {
                    pdf_type = 1;
                }

            }
            Match match3;
            try
            {

                for (int i = 0; i < compressed7.Count; i++)
                {
                    Match match2;
                    Match six_digit;
                    
                    while (!(match2 = Regex.Match(compressed7[i], pattern2)).Success)
                    {
                        i = i + 1;

                    }

                    number_of_subjects = 0;
                    k = i;
                    do
                    {                       
                        subject = Regex.Match(compressed7[k], pattern4);
                        if (subject.Success)
                        {
                            number_of_subjects = number_of_subjects + 1;
                        }
                        k++;
                    }
                    while (!(student_id = Regex.Match(compressed7[k], pattern2)).Success && compressed7[k]!="ENDING" && !(student_id = Regex.Match(compressed7[k], six_digit_pattern)).Success);
                    ArrayList studentlist = new ArrayList();
                    

                    if (match2.Success)
                    {
                        Student student = new Student();
                        student.student_details = compressed7[i];
                        i = i + 1;
                        for (k = 0; k < number_of_subjects; k++)
                        {
                            //while (!(match3 = Regex.Match(compressed7[i], pattern4)).Success)
                            //{
                            //    if((match3 = Regex.Match(compressed7[i], pattern4)).Success)
                            //    {
                            //        break;
                            //    }
                            //    i = i + 1;
                            //}

                            match3 = Regex.Match(compressed7[i], pattern4);
                            if (match3.Success)
                            {
                                Record rec = new Record();
                                rec.subid = compressed7[i];
                                rec.sub_name = compressed7[i + 1];
                                rec.scored_grade = compressed7[i + 2];
                                rec.max_theory_marks = compressed7[i + 3];
                                rec.external_theory_criteria = compressed7[i + 4];
                                rec.scored_external = compressed7[i + 5];
                                rec.max_session_marks = compressed7[i + 6];
                                rec.session_criteria = compressed7[i + 7];
                                rec.session_marks = compressed7[i + 8];
                                rec.max_practical_marks = compressed7[i + 9];
                                rec.practical_criteria = compressed7[i + 10];
                                rec.practical_scored = compressed7[i + 11];
                                rec.max_term_work = compressed7[i + 12];
                                rec.term_work_criteria = compressed7[i + 13];
                                rec.term_work_scored = compressed7[i + 14];
                                rec.total_subject_marks = compressed7[i + 15];
                                rec.minimum_marks_required = compressed7[i + 16];
                                rec.total_marks_scored = compressed7[i + 17];
                                i = i + 18;
                                student.record.Add(rec);

                            }
                            else
                            {
                                i = i + 1;
                                k = k - 1;
                             

                            }




                        }
                        while(compressed7[i]!="SPI")
                        {
                            i = i + 1;
                        }
                        if (pdf_type == 1)
                        {
                            
                            student.spi = compressed7[i + 8];
                            student.cpi = compressed7[i + 10];
                            student.Overall_Marks = compressed7[i + 11];
                            student.Total_Marks_Obtained = compressed7[i + 12];
                            student.Result = compressed7[i + 6];
                            student_list.Add(student);
                            //i = i + 6;
                            i = i + 12;

                        }
                        else if (pdf_type == 0)
                        {
                            student.spi = compressed7[i + 1];
                            student.cpi = compressed7[i + 3];
                            student.Overall_Marks = compressed7[i + 4];
                            student.Total_Marks_Obtained = compressed7[i + 5];
                            student.Result = compressed7[i + 6];
                            student_list.Add(student);
                            i = i + 6;

                        }


                    }

                }

                

            }
            catch(Exception e)
            {
                //File.AppendAllText(path5, e.ToString(), Encoding.UTF8);
            }
            return student_list;

        }



        private string GetTextFromPDF(string fileName)
        {
            StringBuilder text = new StringBuilder();

            using (PdfReader reader = new PdfReader(fileName))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return text.ToString();
        }

        private string[] ConvertToWords(string allLines)
        {
            allWords = null;
            allWords = Regex.Split(allLines, "\\s");
            return allWords;
        }

        private ArrayList allfilters(String input_file, String processing_location)
        {
            allLines = GetTextFromPDF(input_file);
            allWords = ConvertToWords(allLines);
            myCollection = removing_blank_words(allWords, processing_location);
            compressed = removing_waste_words(myCollection, processing_location);
            compressed2 = merging_name_roll(compressed, processing_location);
            compressed4 = merging_subject_name(compressed2, processing_location);
            compressed5 = adding_dash(compressed4, processing_location);
            compressed6 = next_row_name_revision(compressed5, processing_location);
            compressed7 = delete_duplicate_name(compressed6, processing_location);
            studentlist = make_student_objects_from_stringlist(compressed7);//static subject data
            allstudentdetails(compressed7, processing_location);

            return studentlist;

        }
        public void GenerateExcelFiles(String file_output_loc, ArrayList studentobjects)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Workbook worKbooK1;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;
            Microsoft.Office.Interop.Excel.Worksheet worksheet1;
            //Microsoft.Office.Interop.Excel.Range celLrangE;
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            worKbooK = excel.Workbooks.Add(Type.Missing);
            worKbooK1 = excel.Workbooks.Add(Type.Missing);
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
            worksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK1.ActiveSheet;

            worksheet.Name = "Student Rec(Abstract)";
            worksheet1.Name = "Student Subject Records";

            worksheet.Cells[1, 1] = "Student id";
            worksheet.Cells[1, 2] = "spi";
            worksheet.Cells[1, 3] = "cpi";
            worksheet.Cells[1, 4] = "Overall Marks";
            worksheet.Cells[1, 5] = "Total Marks";
            worksheet.Cells[1, 6] = "Result";

            worksheet1.Cells[1, 1] = "student_details";
            worksheet1.Cells[1, 2] = "Subject Code";
            worksheet1.Cells[1, 3] = "Subject Name";
            worksheet1.Cells[1, 4] = "Max Theory Marks";
            worksheet1.Cells[1, 5] = "external_theory_criteria";
            worksheet1.Cells[1, 6] = "scored_external";
            worksheet1.Cells[1, 7] = "max_session_mark";
            worksheet1.Cells[1, 8] = "session_criteria";
            worksheet1.Cells[1, 9] = "session_marks";
            worksheet1.Cells[1, 10] = "max_practical_marks";
            worksheet1.Cells[1, 11] = "practical_criteria";
            worksheet1.Cells[1, 12] = "practical_scored";
            worksheet1.Cells[1, 13] = "max_term_work";
            worksheet1.Cells[1, 14] = "term_work_criteria";
            worksheet1.Cells[1, 15] = "term_work_scored";
            worksheet1.Cells[1, 16] = "total_subject_mark";
            worksheet1.Cells[1, 17] = "minimum_marks_required";
            worksheet1.Cells[1, 18] = "total_marks_scored";
            worksheet1.Cells[1, 19] = "scored_grade";

            

            int row = 2;
            foreach (Student s in studentobjects)
            {
                Match match = Regex.Match(s.student_details, pattern6);

                worksheet.Cells[row, 1] = match.Value;
                worksheet.Cells[row, 2] = s.spi;
                worksheet.Cells[row, 3] = s.cpi;
                worksheet.Cells[row, 4] = s.Overall_Marks;
                worksheet.Cells[row, 5] = s.Total_Marks_Obtained;
                worksheet.Cells[row, 6] = s.Result;

                row = row + 1;
            }

            row = 2;

            foreach (Student s in studentobjects)
            {
                for (int i = 0; i < s.record.Count; i++)
                {
                    Match match = Regex.Match(s.student_details, pattern6);

                    worksheet1.Cells[row, 1] = match.Value;
                    worksheet1.Cells[row, 2] = s.record[i].subid;
                    worksheet1.Cells[row, 3] = s.record[i].sub_name;
                    worksheet1.Cells[row, 4] = s.record[i].max_theory_marks;
                    worksheet1.Cells[row, 5] = s.record[i].external_theory_criteria;
                    worksheet1.Cells[row, 6] = s.record[i].scored_external;
                    worksheet1.Cells[row, 7] = s.record[i].max_session_marks;
                    worksheet1.Cells[row, 8] = s.record[i].session_criteria;
                    worksheet1.Cells[row, 9] = s.record[i].session_marks;
                    worksheet1.Cells[row, 10] = s.record[i].max_practical_marks;
                    worksheet1.Cells[row, 11] = "[" + s.record[i].practical_criteria + "]";
                    worksheet1.Cells[row, 12] = s.record[i].practical_scored;
                    worksheet1.Cells[row, 13] = s.record[i].max_term_work;
                    worksheet1.Cells[row, 14] = "[" + s.record[i].term_work_criteria + "]";
                    worksheet1.Cells[row, 15] = s.record[i].term_work_scored;
                    worksheet1.Cells[row, 16] = s.record[i].total_subject_marks;
                    worksheet1.Cells[row, 17] = s.record[i].minimum_marks_required;
                    worksheet1.Cells[row, 18] = s.record[i].total_marks_scored;
                    worksheet1.Cells[row, 19] = s.record[i].scored_grade;

                    row = row + 1;
                }
            }


            worKbooK.SaveAs(file_output_loc + @"\student_abs.xls");
            worKbooK1.SaveAs(file_output_loc + @"\student_allsub.xls");
            worKbooK.Close();
            worKbooK1.Close();
            excel.Quit();
            //con.Close();
            Console.Write("Student Results has been added to Database \n");

        }


        public void start(string input_file,string output_location,string processing_location)
        {
            ArrayList studentobjects = new ArrayList();
        
            studentobjects = this.allfilters(input_file, processing_location);
            try
            {
                GenerateExcelFiles(output_location, studentobjects);

            }
            catch (Exception e)
            {
             
            }





        }
    }
}
