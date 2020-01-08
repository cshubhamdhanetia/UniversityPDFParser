using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfParserForm
{
    class Student
    {
        public String student_details { get; set; }
        public List<Record> record { get; set; }
        //public ArrayList record { get; set; }


        public String spi { get; set; }
        public String cpi { get; set; }
        public String Overall_Marks { get; set; }
        public String Total_Marks_Obtained { get; set; }
        public String Result { get; set; }

        public Student()
        {
            record = new List<Record>();
        }
    }
}
