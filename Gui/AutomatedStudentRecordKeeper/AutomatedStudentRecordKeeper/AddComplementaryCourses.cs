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
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Principal;

namespace AutomatedStudentRecordKeeper
{
    public partial class AddComplementaryCourses : Form
    {
        string selectedInputFile;
        NpgsqlCommand cmd = new NpgsqlCommand();

        public AddComplementaryCourses()
        {
            InitializeComponent();
            //Adds years to dropdown box
            int tempyear = DateTime.Now.Year;
            for (int i = 6; i >= -1; i--)
            {
                this.yeardropbox.Items.Add((tempyear - i).ToString() + "/" + (tempyear - i + 1).ToString());
            }
        }

        private void submittable_Click_1(object sender, EventArgs e)
        {
            CourseTable.Hide();
            int count = 0;
            //connections string
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                //Checks for Valid Data
                if (yeardropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Year");
                }
                else if (listbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a List");
                }
                else
                {
                    for (int j = 0; j < this.CourseTable.RowCount; j++)
                    {
                        if (string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(0, j).Text) ||
                            string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(1, j).Text) ||
                            string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(2, j).Text) ||
                            string.IsNullOrWhiteSpace(CourseTable.GetControlFromPosition(3, j).Text))
                        {

                        }
                        else if (CourseTable.GetControlFromPosition(0, j).Text.Length != 4 || CourseTable.GetControlFromPosition(1, j).Text.Length != 4)
                        {

                        }
                        else
                        {
                            NpgsqlDataReader reader;
                            NpgsqlCommand cmd;
                            string checkifexists = "False";
                            //query to check if course already exists
                            cmd = new NpgsqlCommand("select exists(select true from courses where coursesubject =  :sub and coursenumber = :num and yearsection = :year)", conn);
                            cmd.Parameters.Add(new NpgsqlParameter("sub", CourseTable.GetControlFromPosition(0, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("num", CourseTable.GetControlFromPosition(1, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("year", yeardropbox.Text));
                            reader = cmd.ExecuteReader();
                            while (reader.Read())
                            {
                                checkifexists = reader[0].ToString();
                            }
                            cmd.Cancel();
                            reader.Close();
                            if (checkifexists == "False")
                            {
                                string checktype = "";
                                //detrmine what list to add to
                                if (listbox.Text == "A")
                                {
                                    checktype = "compa";
                                }
                                else
                                {
                                    checktype = "compb";
                                }
                                //query to add complementary course to database
                                cmd = new NpgsqlCommand("insert into courses values(:sub, :num, 'complementary' ,:name, :cred, NULL,:yearsec,:entyear,:type)", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("sub", CourseTable.GetControlFromPosition(0, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("num", CourseTable.GetControlFromPosition(1, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("name", CourseTable.GetControlFromPosition(2, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("cred", double.Parse(CourseTable.GetControlFromPosition(3, j).Text)));
                                cmd.Parameters.Add(new NpgsqlParameter("yearsec", yeardropbox.Text));
                                cmd.Parameters.Add(new NpgsqlParameter("entyear", int.Parse(yeardropbox.Text.Substring(0, 4))));
                                cmd.Parameters.Add(new NpgsqlParameter("type", checktype));
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (NpgsqlException ex)
                                {

                                }
                            }
                            //clear table after entered
                            CourseTable.GetControlFromPosition(0, j).Text = "";
                            CourseTable.GetControlFromPosition(1, j).Text = "";
                            CourseTable.GetControlFromPosition(2, j).Text = "";
                            CourseTable.GetControlFromPosition(3, j).Text = "";
                            count++;
                        }
                    }
                    conn.Close();
                    //message displaying successful queries
                    MessageBox.Show(count.ToString() + " rows added to table, check formating if any field not cleared");
                }
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
            CourseTable.Show();
        }
        //key press event to only allow numbers
        private void richTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }

        //key press event to only allow upper case letters
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

        private void importCSV_Click(object sender, EventArgs e)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                if (yeardropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Year");
                }
                else
                {
                    OpenFileDialog choofdlog = new OpenFileDialog(); //opens file viewer
                    choofdlog.Filter = "csv Files (*.csv*)|*.csv"; //only shows csv files
                    choofdlog.Title = "Select a CSV File";
                    choofdlog.FilterIndex = 1;
                    choofdlog.Multiselect = false; //one file at a time

                    if (choofdlog.ShowDialog() == DialogResult.OK)
                    {
                        selectedInputFile = choofdlog.FileName; //sets path
                        string csvFile = System.IO.File.ReadAllText(selectedInputFile); //reads from path

                        //clean file here//
                        //Change names to coursecode
                        csvFile = Regex.Replace(csvFile, @"\s(engineering)(?=(\s[0-9]{4}))", "ENGI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"\s(english)(?=(\s[0-9]{4}))", "ENGL", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(sociology)(?=(\s[0-9]{4}))", "SOCI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(economics)(?=(\s[0-9]{4}))", "ECON", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(anthropology)(?=(\s[0-9]{4}))", "ANTH", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(criminology)(?=(\s[0-9]{4}))", "CRIM", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(geography)(?=(\s[0-9]{4}))", "GEOG", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(history)(?=(\s[0-9]{4}))", "HIST", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(indigenous learning)(?=(\s[0-9]{4}))", "INDI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(finnish)(?=(\s[0-9]{4}))", "FINN", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(french)(?=(\s[0-9]{4}))", "FREN", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(italian)(?=(\s[0-9]{4}))", "ITAL", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(linguistics)(?=(\s[0-9]{4}))", "LING", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(music)(?=(\s[0-9]{4}))", "MUSI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(philosophy)(?=(\s[0-9]{4}))", "PHIL", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(political science)(?=(\s[0-9]{4}))", "POLI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(psychology)(?=(\s[0-9]{4}))", "PSYC", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(visual arts)(?=(\s[0-9]{4}))", "VISU", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(women's studies)(?=(\s[0-9]{4}))", "WOME", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(business)(?=(\s[0-9]{4}))", "BUSI", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(german)(?=(\s[0-9]{4}))", "GERM", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(language)(?=(\s[0-9]{4}))", "LANG", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(ojibwe)(?=(\s[0-9]{4}))", "OJIB", RegexOptions.IgnoreCase);
                        csvFile = Regex.Replace(csvFile, @"(spanish)(?=(\s[0-9]{4}))", "SPAN", RegexOptions.IgnoreCase);


                        csvFile = Regex.Replace(csvFile, @"\(Prerequisite:(.*?)\""", "\n");

                        csvFile = Regex.Replace(csvFile, @" \(.*?\w{2}\)", string.Empty, RegexOptions.Multiline); //removes most info in brackets

                        csvFile = Regex.Replace(csvFile, @"(?<=([A-Z]{4}))\s", "~"); //setting up delimiter
                        csvFile = Regex.Replace(csvFile, @"\b-\s", " - "); //setting up delimiter
                        csvFile = Regex.Replace(csvFile, @"\s-\s", "~comp~"); //setting up delimiter with coursesection
                        csvFile = Regex.Replace(csvFile, @"(?<=\d):\s", "~comp~"); //setting up delimiter with coursesection
                        csvFile = Regex.Replace(csvFile, @"\""�", string.Empty); //removes bullets
                        csvFile = Regex.Replace(csvFile, @"\s�\s", "~comp~"); //setting up delimiter with coursesection for unicode character similar to hyphen
                        csvFile = Regex.Replace(csvFile, @"OR", "\n"); //rare case of two courses on same line
                        csvFile = Regex.Replace(csvFile, @"\(credit weight 1.0\)|\(credit weight 1.0 \)", "~1.0"); //setting up delimiter
                        csvFile = Regex.Replace(csvFile, @" \(.*?\n", "\n");
                        csvFile = Regex.Replace(csvFile, @"\""", string.Empty); //removes remaining quotations

                        string[] lists = csvFile.Split(new string[] { "List B: Other Complementary Studies Courses" }, StringSplitOptions.RemoveEmptyEntries);

                        string compCoursesFile = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "complimentaryCourses.txt");
                        System.IO.File.Delete(compCoursesFile);
                        using (System.IO.File.Create(compCoursesFile)) { }

                        //sets permissions of file
                        FileSecurity access = System.IO.File.GetAccessControl(compCoursesFile);
                        SecurityIdentifier everyone = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
                        access.AddAccessRule(new FileSystemAccessRule(everyone, FileSystemRights.ReadAndExecute, AccessControlType.Allow));
                        System.IO.File.SetAccessControl(compCoursesFile, access);

                        for (int j = 0; j < lists.Length; j++)
                        {
                            Console.WriteLine("Loop {0}", j);
                            //creating list to prepare data for entering into dB
                            var courseCode = Regex.Matches(lists[j], @"[A-Z]{4}~[0-9]{4}.*?[\n|\r]");
                            var courseCodeList = courseCode.Cast<Match>().Select(match => match.Value).ToList();

                            for (int i = 0; i < courseCodeList.Count; i++)
                            {
                                if (!courseCodeList[i].Contains("1.0"))
                                {
                                    courseCodeList[i] += "~0.5~" + int.Parse(yeardropbox.Text.Substring(0, 4));
                                    courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    if (j == 0)
                                    {
                                        courseCodeList[i] += "~compa";
                                        courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    }
                                    else if (j == 1)
                                    {
                                        courseCodeList[i] += "~compb";
                                        courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    }
                                }
                                else
                                {
                                    courseCodeList[i] += "~" + int.Parse(yeardropbox.Text.Substring(0, 4));
                                    courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    if (j == 0)
                                    {
                                        courseCodeList[i] += "~compa";
                                        courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    }
                                    else if (j == 1)
                                    {
                                        courseCodeList[i] += "~compb";
                                        courseCodeList[i] = Regex.Replace(courseCodeList[i], @"[\n\r]+", string.Empty);
                                    }
                                }
                            }

                            //creates file for bulk insert data into dB
                            System.IO.File.AppendAllLines(compCoursesFile, courseCodeList.Distinct().ToArray()); //writes to text file

                        }
                        try
                        {
                            cmd = new NpgsqlCommand("COPY courses (coursesubject, coursenumber, coursesection, coursename, credits, yearused, type) FROM "
                            + "'" + compCoursesFile + "'" + " DELIMITER '~' CSV", conn);
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            goto ERROR;
                        }

                        var result = MessageBox.Show("Importing of Complementary courses to the database was successful. Would you like to import another file?", "SUCCESS",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                            this.Activate();
                        if (result == DialogResult.No)
                            this.Close();
                        goto END;
                    ERROR:
                        MessageBox.Show("Error Adding Complementary Courses to DataBase", "ERROR");
                        this.Close();
                    END:
                        Debug.WriteLine("DONE");

                    }
                    else
                        selectedInputFile = string.Empty;
                }
            }
        }

        private void AddComplementaryCourses_Load(object sender, EventArgs e)
        {

        }

        private void yeardropbox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}