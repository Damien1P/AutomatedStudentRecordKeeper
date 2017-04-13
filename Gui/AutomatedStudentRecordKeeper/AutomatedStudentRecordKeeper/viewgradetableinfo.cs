using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedStudentRecordKeeper
{
    class viewgradetableinfo
    {
        private string subject;
        private string number;
        private string name;
        private string grade;
        private string credits;

        public string Subject
        {
            get
            {
                return subject;
            }

            set
            {
                subject = value;
            }
        }
        public string Number
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
        public string Name
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
        public string Grade
        {
            get
            {
                return grade;
            }

            set
            {
                grade = value;
            }
        }
        public string Credits
        {
            get
            {
                return credits;
            }

            set
            {
                credits = value;
            }
        }
    }
}
