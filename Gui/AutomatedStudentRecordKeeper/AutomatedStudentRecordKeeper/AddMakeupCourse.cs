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
    public partial class AddMakeupCourse : Form
    {
        public AddMakeupCourse()
        {
            InitializeComponent();
        }

        private void studentnumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }

        private void richTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (Char.IsDigit(ch))
            {
                e.Handled = true;
            }
            else
            {
                e.KeyChar = Char.ToUpper(ch);
            }
        }

        private void submittable_Click(object sender, EventArgs e)
        {
            int count = 0;
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                NpgsqlCommand cmd = new NpgsqlCommand("select exists (select true from student where studentnumber = '" + studentnumber.Text + "')", conn);
                string checknum = "False";
                NpgsqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    checknum = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                if (studentnumber.Text.Length!=7)
                {
                    MessageBox.Show("Please enter valid student number");
                }
                else if (checknum == "False")
                {
                    MessageBox.Show("Student Doesnt exists");
                }
                else
                {
                    CourseTable.Hide();
                    for (int j = 0; j < this.CourseTable.RowCount; j++)
                    {
                        if (string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(0, j).Text) ||
                            string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(1, j).Text))
                        {

                        }
                        else if(CourseTable.GetControlFromPosition(0, j).Text.Length!=4 || CourseTable.GetControlFromPosition(1, j).Text.Length != 4)
                        {

                        }
                        else    
                        {
                            string studyear="";
                            cmd = new NpgsqlCommand("select year from student where studentnumber = '" + studentnumber.Text + "'", conn);
                            reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                studyear = reader[0].ToString();
                            }
                            cmd.Cancel();
                            reader.Close();

                            cmd = new NpgsqlCommand("insert into courses values(:sub, :num, NULL,:name,:cred,NULL,NULL,:year,'makeup')", conn);
                            cmd.Parameters.Add(new NpgsqlParameter("sub", CourseTable.GetControlFromPosition(0, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("num", CourseTable.GetControlFromPosition(1, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("name", CourseTable.GetControlFromPosition(2, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("cred", 0.5));
                            cmd.Parameters.Add(new NpgsqlParameter("year", int.Parse(studyear)));
                            try
                            {
                                cmd.ExecuteNonQuery();
                            }
                            catch (NpgsqlException ex)
                            {

                            }
                            cmd.Cancel();
                            count++;
                            CourseTable.GetControlFromPosition(0, j).Text = "";
                            CourseTable.GetControlFromPosition(1, j).Text = "";
                            CourseTable.GetControlFromPosition(2, j).Text = "";
                        }
                    }
                    MessageBox.Show(count.ToString()+" rows added, if row not cleared check formating");
                    CourseTable.Show();
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
        }
    }
}

