using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedStudentRecordKeeper
{
    class viewtabletinfo
    {
        private string subject;
        private string number;
        private string title;
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
        public string Title
        {
            get
            {
                return title;
            }

            set
            {
                title = value;
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
