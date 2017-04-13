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
    public partial class ViewStudents : Form
    {
        List<studentinfo> students = new List<studentinfo>();
        public ViewStudents()
        {
            InitializeComponent();
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {

                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                cmd = new NpgsqlCommand("select name, studentnumber from student ", conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    students.Add(new studentinfo { StudentName = reader[0].ToString(), StudentNumber = reader[1].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                studenttable.DataSource = students;
                conn.Close();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
        }

        private void removebutton_Click(object sender, EventArgs e)
        {

            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                if (removetext.Text.Length != 7)
                {
                    MessageBox.Show("Please enter valid student number");
                }
                else
                {
                    NpgsqlCommand cmd;
                    cmd = new NpgsqlCommand("delete from student where studentid = '" + removetext.Text + "'", conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Student removed if existed");
                }
                    conn.Close();
                
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }

        }

        private void removetext_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }
    }
    }
