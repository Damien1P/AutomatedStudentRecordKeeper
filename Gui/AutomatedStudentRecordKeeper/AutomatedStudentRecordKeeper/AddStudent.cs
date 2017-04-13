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
    public partial class AddStudent : Form
    {
        public AddStudent()
        {
            InitializeComponent();
        }

        private void TransferStudent_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void AddRowMakeup_Click(object sender, EventArgs e)
        {
            RowStyle temp = MakeupCourseTable.RowStyles[MakeupCourseTable.RowCount - 1];
            
            MakeupCourseTable.RowCount++;
            MakeupCourseTable.RowStyles.Add(new RowStyle(temp.SizeType, temp.Height));
            MakeupCourseTable.Controls.Add(new RichTextBox() { Width = 46, Height = 21,MaxLength = 4 }, 0, MakeupCourseTable.RowCount - 1);
            MakeupCourseTable.GetControlFromPosition(0,MakeupCourseTable.RowCount - 1).KeyPress+= new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            MakeupCourseTable.Controls.Add(new RichTextBox() { Width = 53, Height = 21, MaxLength = 4 }, 1, MakeupCourseTable.RowCount - 1);
            MakeupCourseTable.GetControlFromPosition(1, MakeupCourseTable.RowCount - 1).KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            MakeupCourseTable.Controls.Add(new RichTextBox() { Width = 233, Height = 21 }, 2, MakeupCourseTable.RowCount - 1);
        }

        private void AddRowCredits_Click(object sender, EventArgs e)
        {
            RowStyle temp = CourseCredTable.RowStyles[CourseCredTable.RowCount - 1];
            CourseCredTable.RowCount++;
            CourseCredTable.RowStyles.Add(new RowStyle(temp.SizeType, temp.Height));
            CourseCredTable.Controls.Add(new RichTextBox() { Width = 46, Height = 21 }, 0, CourseCredTable.RowCount - 1);
            CourseCredTable.GetControlFromPosition(0, CourseCredTable.RowCount - 1).KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            CourseCredTable.Controls.Add(new RichTextBox() { Width = 53, Height = 21 }, 1, CourseCredTable.RowCount - 1);
            CourseCredTable.GetControlFromPosition(1, CourseCredTable.RowCount - 1).KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            CourseCredTable.Controls.Add(new RichTextBox() { Width = 233, Height = 21 }, 2, CourseCredTable.RowCount - 1);
        }

        private void tableLayoutPanel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void AddStudentButton_Click(object sender, EventArgs e)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {

                if (StudentNumberBox.Text.Length != 7)
                {
                    MessageBox.Show("Please input valid student number");
                }
                else if (Yeardropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select year");
                }
                else if (string.IsNullOrWhiteSpace(Namebox.Text))
                {
                    MessageBox.Show("Please input a name");
                }
                else
                {
                    NpgsqlCommand cmd = new NpgsqlCommand("select exists (select true from student where studentnumber = '" + StudentNumberBox.Text + "')", conn);
                    string checknum = "False";
                    NpgsqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        checknum = reader[0].ToString();
                    }
                    cmd.Cancel();
                    reader.Close();
                    if (checknum != "False")
                    {
                        MessageBox.Show("Student Already exists");
                    }
                    else if (checkcredit())
                    {
                        MessageBox.Show("Please check formating of rows of Course Credit Table");
                    }
                    else if (checkmakeup())
                    {
                        MessageBox.Show("Please check formating of rows of Make Up Course Table");
                    }
                    else
                    {

                        cmd = new NpgsqlCommand("insert into student values(:name, :num, :prevsc, :prevprog, :yrlvl,:yrsec ,:entyear)", conn);
                        cmd.Parameters.Add(new NpgsqlParameter("name", Namebox.Text));
                        cmd.Parameters.Add(new NpgsqlParameter("num", StudentNumberBox.Text));
                        cmd.Parameters.Add(new NpgsqlParameter("prevsc", PrevSchoolBox.Text));
                        cmd.Parameters.Add(new NpgsqlParameter("prevprog", PrevProgramBox.Text));
                        cmd.Parameters.Add(new NpgsqlParameter("yrlvl", int.Parse(Yeardropbox.Text)));
                        cmd.Parameters.Add(new NpgsqlParameter("yrsec", (DateTime.Now.Year -1 ).ToString() + "/" + (DateTime.Now.Year).ToString()));
                        cmd.Parameters.Add(new NpgsqlParameter("entyear", DateTime.Now.Year -1));
                        try
                        {
                            cmd.ExecuteNonQuery();
                        }
                        catch (NpgsqlException ex)
                        {

                        }
                        cmd.Cancel();
                        
                        for (int j = 0; j < this.MakeupCourseTable.RowCount; j++)
                        {
                            if (string.IsNullOrWhiteSpace(MakeupCourseTable.GetControlFromPosition(0, j).Text) ||
                                string.IsNullOrWhiteSpace(MakeupCourseTable.GetControlFromPosition(1, j).Text))
                            {

                            }
                            else
                            {
                                cmd = new NpgsqlCommand("insert into courses values(:sub, :num, NULL,:name,:cred,NULL,NULL,:year,'makeup')", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("sub", MakeupCourseTable.GetControlFromPosition(0, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("num", MakeupCourseTable.GetControlFromPosition(1, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("name", MakeupCourseTable.GetControlFromPosition(2, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("cred", 0.5));
                                cmd.Parameters.Add(new NpgsqlParameter("year", DateTime.Now.Year-1));
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (NpgsqlException ex)
                                {

                                }
                            }
                        }

                        cmd.Cancel();
                        for (int j = 0; j < this.CourseCredTable.RowCount; j++)
                        {
                            if (string.IsNullOrWhiteSpace(CourseCredTable.GetControlFromPosition(0, j).Text) ||
                                string.IsNullOrWhiteSpace(CourseCredTable.GetControlFromPosition(1, j).Text))
                            {

                            }
                            else
                            {
                                cmd = new NpgsqlCommand("insert into grades values(:stnum, :sub, :num, NULL, :yr)", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("stnum", StudentNumberBox.Text));
                                cmd.Parameters.Add(new NpgsqlParameter("sub", CourseCredTable.GetControlFromPosition(0, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("num", CourseCredTable.GetControlFromPosition(1, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("yr", DateTime.Now.Year));
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (NpgsqlException ex)
                                {

                                }
                            }
                        }
                        conn.Close();
                        MessageBox.Show("Student Successfully Added");
                        this.Close();
                    }
                    conn.Close();
                }

            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
            }

        private void Namebox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (Char.IsDigit(ch))
            {
                e.Handled = true;
            }
        }

        private void StudentNumberBox_KeyPress(object sender, KeyPressEventArgs e)
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
        public bool checkmakeup()
        {
            bool check = false;
            for (int i = 0; i < MakeupCourseTable.RowCount; i++)
            {
                if(MakeupCourseTable.GetControlFromPosition(0,i).Text.Length!=4 
                    || MakeupCourseTable.GetControlFromPosition(1, i).Text.Length != 4)
                {
                    if (string.IsNullOrWhiteSpace(MakeupCourseTable.GetControlFromPosition(0, i).Text) && string.IsNullOrWhiteSpace(MakeupCourseTable.GetControlFromPosition(1, i).Text))
                    {

                    }
                    else
                    {
                        check = true;
                    }
                }
            }
            return check;
        }
        public bool checkcredit()
        {
            bool check = false;
            for (int i = 0; i < CourseCredTable.RowCount; i++)
            {
                if (CourseCredTable.GetControlFromPosition(0, i).Text.Length != 4
                    || CourseCredTable.GetControlFromPosition(1, i).Text.Length != 4)
                {
                    if (string.IsNullOrWhiteSpace(CourseCredTable.GetControlFromPosition(0, i).Text) && string.IsNullOrWhiteSpace(CourseCredTable.GetControlFromPosition(1, i).Text))
                    {

                    }
                    else
                    {
                        check = true;
                    }
                }
            }
            return check;
        }
    }
    } 
   

