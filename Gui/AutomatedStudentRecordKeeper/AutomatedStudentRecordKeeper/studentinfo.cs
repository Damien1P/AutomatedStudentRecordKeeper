using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedStudentRecordKeeper
{
    class studentinfo
    {
        private string name;
        private string number;
       
        public string StudentName
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }
        public string StudentNumber
        {
            get
            {
                return number;
            }

            set
            {
                number = value;
            }
        }
    }
}
