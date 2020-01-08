using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfParserForm
{
    class Record
    {
        public String subid { get; set; }
        public String sub_name { get; set; }
        public String max_theory_marks { get; set; }
        public String external_theory_criteria { get; set; }
        public String scored_external { get; set; }
        public String max_session_marks { get; set; }
        public String session_criteria { get; set; }
        public String session_marks { get; set; }
        public String max_practical_marks { get; set; }
        public String practical_criteria { get; set; }
        public String practical_scored { get; set; }
        public String max_term_work { get; set; }
        public String term_work_criteria { get; set; }
        public String term_work_scored { get; set; }
        public String total_subject_marks { get; set; }
        public String minimum_marks_required { get; set; }
        public String total_marks_scored { get; set; }
        public String scored_grade { get; set; }
    }
}
