using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using NpgsqlTypes;

namespace AutomatedStudentRecordKeeper
{
    public partial class MainScreen : Form
    {
        public MainScreen()
        {
            InitializeComponent();
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                NpgsqlCommand cmd;
                //querys to remove old data from database
                cmd = new NpgsqlCommand("delete from student where year <= " +(DateTime.Now.Year - 7).ToString(), conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel();
                cmd = new NpgsqlCommand("delete from courses where yearused <= " + (DateTime.Now.Year - 7).ToString(), conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel();
                conn.Close();

            }
            else
            {
                MessageBox.Show("Connection error to database");
            }

        
    }
        //All click events for each button to access the other windows of the GUI
             private void addstudentbutton_Click(object sender, EventArgs e)
        {
            AddStudent add_student = new AddStudent();
            add_student.ShowDialog();
        }

        private void viewstudentsbutton_Click(object sender, EventArgs e)
        {
            ViewStudents view_student = new ViewStudents();
            view_student.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddStudent add_transfer = new AddStudent();
            add_transfer.ShowDialog();
        }

        private void AddCoursesButton_Click(object sender, EventArgs e)
        {
            AddCourse add_course = new AddCourse();
            add_course.ShowDialog();
        }

        private void ViewCoursesButton_Click(object sender, EventArgs e)
        {
            ViewCourse view_course = new ViewCourse();
            view_course.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddComplementaryCourses add_comp_course = new AddComplementaryCourses();
            add_comp_course.ShowDialog();
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            AddGrade add_grade = new AddGrade();
            add_grade.ShowDialog();
        }

        private void ViewButton_Click(object sender, EventArgs e)
        {
            ViewGrade view_grade = new ViewGrade();
            view_grade.ShowDialog();
        }

        private void addmakeup_Click(object sender, EventArgs e)
        {
            AddMakeupCourse add_makeup = new AddMakeupCourse();
            add_makeup.ShowDialog();
        }
    }
}
