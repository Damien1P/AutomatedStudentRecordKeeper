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
    public partial class ViewGrade : Form
    {
        string studentnumber="";
        bool dataloaded = false;
        string curricyear = "";
        BindingList<viewgradetableinfo> tableviewlist = new BindingList<viewgradetableinfo>();
        public ViewGrade()
        {
            InitializeComponent();
            Fulltable1.Hide();
            Fulltable2.Hide();
            Fulltable3.Hide();
            Fulltable4.Hide();
            Fulltable5.Hide();
            Fulltable6.Hide();


        }
        private void Calculate_Click(object sender, EventArgs e)
        {
            tableviewlist.Clear();
            yearbox.Items.Clear();
            dataloaded = false;
            List<string> years = new List<string>();
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                if (StudentBox.Text.Length != 7)
                {
                    MessageBox.Show("Please enter valid student number");
                }
                else
                {
                    NpgsqlCommand cmd = new NpgsqlCommand("select exists (select true from student where studentnumber = '" + StudentBox.Text + "')", conn);
                    string checknum = "False";
                    NpgsqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        checknum = reader[0].ToString();
                    }
                    cmd.Cancel();
                    reader.Close();
                    if (checknum == "False")
                    {
                        MessageBox.Show("Student Doesn't exist");
                    }
                    else
                    {
                        cmd = new NpgsqlCommand("select year from student where studentnumber = '"+StudentBox.Text+"'", conn);
                        reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            curricyear = reader[0].ToString();
                        }
                        cmd.Cancel();
                        reader.Close();
                        string temp="";
                        waitScreen waitscrn = new waitScreen();
                        waitscrn.Show();
                        Application.DoEvents();
                        cmd = new NpgsqlCommand("select distinct yeartaken from grades where studentnumber = '" + StudentBox.Text + "' order by yeartaken asc", conn);
                        reader = cmd.ExecuteReader();
                        while (reader.Read())
                        {
                            temp = reader[0].ToString();
                            temp = temp + "/" + (int.Parse(temp) + 1).ToString();
                            years.Add(temp);
                            yearbox.Items.Add(temp);
                        }
                        cmd.Cancel();
                        reader.Close();

                        Fulltable1.Hide();
                        Fulltable2.Hide();
                        Fulltable3.Hide();
                        Fulltable4.Hide();
                        Fulltable5.Hide();
                        Fulltable6.Hide();
                        
                        switch (years.Count)
                        {
                            case 1:
                                date1.Text = years[0];
                                Loadtable(table1,date1, SCA1);
                                Fulltable1.Show();
                                break;
                            case 2:
                                date1.Text = years[0];
                                date2.Text = years[1];
                                Loadtable(table1, date1, SCA1);
                                Loadtable(table2, date2, SCA2);
                                Fulltable1.Show();
                                Fulltable2.Show();
                                break;
                            case 3:
                                date1.Text = years[0];
                                date2.Text = years[1];
                                date3.Text = years[2];
                                Loadtable(table1, date1, SCA1);
                                Loadtable(table2, date2, SCA2);
                                Loadtable(table3, date3, SCA3);
                                Fulltable1.Show();
                                Fulltable2.Show();
                                Fulltable3.Show();
                                break;
                            case 4:
                                date1.Text = years[0];
                                date2.Text = years[1];
                                date3.Text = years[2];
                                date4.Text = years[3];
                                Loadtable(table1, date1, SCA1);
                                Loadtable(table2, date2, SCA2);
                                Loadtable(table3, date3, SCA3);
                                Loadtable(table4, date4, SCA4);
                                Fulltable1.Show();
                                Fulltable2.Show();
                                Fulltable3.Show();
                                Fulltable4.Show();
                                break;
                            case 5:
                                date1.Text = years[0];
                                date2.Text = years[1];
                                date3.Text = years[2];
                                date4.Text = years[3];
                                date5.Text = years[4];
                                Loadtable(table1, date1, SCA1);
                                Loadtable(table2, date2, SCA2);
                                Loadtable(table3, date3, SCA3);
                                Loadtable(table4, date4, SCA4);
                                Loadtable(table5, date5, SCA5);
                                Fulltable1.Show();
                                Fulltable2.Show();
                                Fulltable3.Show();
                                Fulltable4.Show();
                                Fulltable5.Show();
                                break;
                            case 6:
                                date1.Text = years[0];
                                date2.Text = years[1];
                                date3.Text = years[2];
                                date4.Text = years[3];
                                date5.Text = years[4];
                                date6.Text = years[5];
                                Loadtable(table1, date1, SCA1);
                                Loadtable(table2, date2, SCA2);
                                Loadtable(table3, date3, SCA3);
                                Loadtable(table4, date4, SCA4);
                                Loadtable(table5, date5, SCA5);
                                Loadtable(table6, date6, SCA6);
                                Fulltable1.Show();
                                Fulltable2.Show();
                                Fulltable3.Show();
                                Fulltable4.Show();
                                Fulltable5.Show();
                                Fulltable6.Show();
                                break;
                            default:
                                MessageBox.Show("Student Has No Grades Yet");
                                break;
                        }

                        dataloaded = true;
                        studentnumber = StudentBox.Text;
                        conn.Close();
                        waitscrn.Hide();
                    }
                }
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }

        }
        public void Loadtable(TableLayoutPanel table,Control date, Control SCA)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '"+StudentBox.Text+"' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and ((c.type ='curric' and c.yearused = "+curricyear+" and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(0, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb') and c.yearused <= :yearnum and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(0, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select (sum(final.grade*final.credits)/sum(final.credits))::int from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,n.grade,c.credits from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken = :yearnum and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(1, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select (sum(final.grade * final.credits)/sum(final.credits))::int from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,n.grade,c.credits from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and c.coursenumber = n.coursenumber and c.coursesubject = n.coursesubject and n.yeartaken = :yearnum and ((c.type ='curric' and c.yearused = "+curricyear+" and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb')  and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(1, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select (sum(final.grade*final.credits)/sum(final.credits))::int from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,n.grade,c.credits from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(2, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select (sum(final.grade * final.credits)/sum(final.credits))::int from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,n.grade,c.credits from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and c.coursenumber = n.coursenumber and c.coursesubject = n.coursesubject and n.yeartaken <= :yearnum and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb')  and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(2, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 60 and n.grade >= 50 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(3, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 60 and n.grade >= 50 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb') and c.yearused <= :yearnum and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(3, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 40 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(4, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 40 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb') and c.yearused <= :yearnum and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(4, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 50 and n.grade >= 40 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 1 or c.yearlevel = 2 )) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(5, 0).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select count(*) from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken <= :yearnum and n.grade < 50 and n.grade >= 40 and ((c.type ='curric' and c.yearused = " + curricyear + " and (c.yearlevel = 3 or c.yearlevel = 4 )) or ((c.type = 'compa' or c.type = 'compb') and c.yearused <= :yearnum and c.yearused = n.yeartaken)) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    table.GetControlFromPosition(5, 1).Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();

                cmd = new NpgsqlCommand("select (sum(final.grade*final.credits)/sum(final.credits))::int from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,n.grade,c.credits from grades as n, courses as c where n.studentnumber = '" + StudentBox.Text + "' and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken = :yearnum and ((c.type ='curric' and c.yearused = " + curricyear + ") or ((c.type = 'compa' or c.type = 'compb')  and c.yearused = n.yeartaken) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(date.Text.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    SCA.Text = reader[0].ToString();
                }
                cmd.Cancel();
                reader.Close();
                conn.Close();
            
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }

        }
        
        public void Loadtablebyyear(string yearsection)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {

                
                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                cmd = new NpgsqlCommand("select final.coursesubject, final.coursenumber,final.coursename,final.grade,final.credits from (select distinct on (n.coursesubject,n.coursenumber) n.coursesubject,n.coursenumber,c.coursename,n.grade,c.credits from grades as n, courses as c where n.studentnumber = :snum and n.coursenumber = c.coursenumber and n.coursesubject = c.coursesubject and n.yeartaken = :yearnum and ((c.type ='curric' and c.yearused = "+curricyear+") or ((c.type = 'compa' or c.type = 'compb')  and c.yearused = n.yeartaken) or (c.type = 'makeup')) order by n.coursesubject, n.coursenumber, n.yeartaken desc )as final union select g.coursesubject,g.coursenumber,NULL,g.grade,NULL from grades as g where g.yeartaken = :yearnum and g.coursenumber not in (select coursenumber from courses)", conn);
                cmd.Parameters.Add(new NpgsqlParameter("snum", studentnumber));
                cmd.Parameters.Add(new NpgsqlParameter("yearnum", int.Parse(yearsection.Substring(0, 4))));
                reader = cmd.ExecuteReader();
                tableviewlist.Clear();
                while (reader.Read())
                {
                    tableviewlist.Add(new viewgradetableinfo { Subject = reader[0].ToString(), Number = reader[1].ToString(), Name = reader[2].ToString(),Grade = reader[3].ToString(), Credits = reader[4].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                gradetable.DataSource = tableviewlist;
                gradetable.Refresh();


                conn.Close();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }

        }

        private void CourseTable_Paint(object sender, PaintEventArgs e)
        {

        }

        private void yearbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataloaded == true)
            {
                Loadtablebyyear(yearbox.Text);
            }
        }
        
        private void StudentBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }
    }
  
}

