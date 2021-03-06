﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;
using Npgsql;
using NpgsqlTypes;
namespace AutomatedStudentRecordKeeper
{
    public partial class ViewCourse : Form
    {
        //Lists to store data from database
        BindingList<viewtabletinfo> falllist1 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> falllist2 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> falllist3 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> falllist4 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> wintlist1 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> wintlist2 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> wintlist3 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> wintlist4 = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> compalist = new BindingList<viewtabletinfo>();
        BindingList<viewtabletinfo> compblist = new BindingList<viewtabletinfo>();
        public ViewCourse()
        {
            InitializeComponent();
            //connection string
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                //gets all course years and adds them to  dropdownboxs
                cmd = new NpgsqlCommand("select distinct yearused from courses where type != 'makeup' order by yearused asc", conn);
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    this.yeardropbox.Items.Add(reader[0].ToString());
                    this.removeyear.Items.Add(reader[0].ToString());
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
        private void loaddatatable(DataGridView falltable, DataGridView winttable, BindingList<viewtabletinfo> list1, BindingList<viewtabletinfo> list2, int yearlevel, int year)
        {
            //connection screen
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                list1.Clear();
                list2.Clear();

                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                //gets all Fall course data for a year and yearlevel
                cmd = new NpgsqlCommand("select coursesubject,coursenumber,coursename,credits FROM courses where coursesection = 'Fall' and yearlevel = :yrlvl and yearused = :year and type = 'curric'", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yrlvl", yearlevel));
                cmd.Parameters.Add(new NpgsqlParameter("year", year));
                reader = cmd.ExecuteReader();
                //adds data to List
                while (reader.Read())
                {
                    list1.Add(new viewtabletinfo { Subject = reader[0].ToString(), Number = reader[1].ToString(), Title = reader[2].ToString(), Credits = reader[3].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                //Puts Fall data into table
                falltable.DataSource = list1;
                //Gets all Winter course data for a year and year level
                cmd = new NpgsqlCommand("select coursesubject,coursenumber,coursename,credits FROM courses where coursesection = 'Winter' and yearlevel = :yrlvl and yearused = :year and type = 'curric'", conn);
                cmd.Parameters.Add(new NpgsqlParameter("yrlvl", yearlevel));
                cmd.Parameters.Add(new NpgsqlParameter("year", year));
                reader = cmd.ExecuteReader();
                //adds data to list
                while (reader.Read())
                {
                    list2.Add(new viewtabletinfo { Subject = reader[0].ToString(), Number = reader[1].ToString(), Title = reader[2].ToString(), Credits = reader[3].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                //puts winter data into table
                winttable.DataSource = list2;
                conn.Close();
                //keep tables up to date
                falltable.Refresh();
                winttable.Refresh();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
        }
        private void loadcomptable(int year)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                compalist.Clear();
                compblist.Clear();
                NpgsqlCommand cmd;
                NpgsqlDataReader reader;
                // gets list A courses
                cmd = new NpgsqlCommand("select coursesubject,coursenumber,coursename,credits FROM courses where yearused = :year and type = 'compa'", conn);
                cmd.Parameters.Add(new NpgsqlParameter("year", year));
                reader = cmd.ExecuteReader();
                //Adds courses to List A
                while (reader.Read())
                {
                    compalist.Add(new viewtabletinfo { Subject = reader[0].ToString(), Number = reader[1].ToString(), Title = reader[2].ToString(), Credits = reader[3].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                //Adds course to list a table
                listatable.DataSource = compalist;
                // gets list b courses
                cmd = new NpgsqlCommand("select coursesubject,coursenumber,coursename,credits FROM courses where yearused = :year and type = 'compb'", conn);
                cmd.Parameters.Add(new NpgsqlParameter("year", year));
                reader = cmd.ExecuteReader();
                //Adds course to List B
                while (reader.Read())
                {
                    compblist.Add(new viewtabletinfo { Subject = reader[0].ToString(), Number = reader[1].ToString(), Title = reader[2].ToString(), Credits = reader[3].ToString() });
                }
                cmd.Cancel();
                reader.Close();
                //Adds courses to list b table
                listbtable.DataSource = compblist;
                conn.Close();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
        }
        //Loads all tables when a year is selected
        private void yeardropbox_SelectedIndexChanged(object sender, EventArgs e)
        {
            waitScreen waitscrn = new waitScreen();
            waitscrn.Show();
            Application.DoEvents();
            loaddatatable(fall1table, winter1table, falllist1, wintlist1, 1, int.Parse(yeardropbox.Text));
            loaddatatable(fall2table, winter2table, falllist2, wintlist2, 2, int.Parse(yeardropbox.Text));
            loaddatatable(fall3table, winter3table, falllist3, wintlist3, 3, int.Parse(yeardropbox.Text));
            loaddatatable(fall4table, winter4table, falllist4, wintlist4, 4, int.Parse(yeardropbox.Text));
            loadcomptable(int.Parse(yeardropbox.Text));
            waitscrn.Hide();

        }

        private void removebutton_Click(object sender, EventArgs e)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                NpgsqlCommand cmd;
                //query for removing courses from database
                cmd = new NpgsqlCommand("DELETE FROM courses WHERE coursesubject = :sub and coursenumber = :num and yearused = :year", conn);
                cmd.Parameters.Add(new NpgsqlParameter("sub", coursesubject.Text));
                cmd.Parameters.Add(new NpgsqlParameter("num", coursenumber.Text));
                cmd.Parameters.Add(new NpgsqlParameter("year", int.Parse(removeyear.Text)));
                cmd.ExecuteNonQuery();
                cmd.Cancel();
                conn.Close();
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
        }
        //keypress to only allow numbers
        private void coursenumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }
    }
}
