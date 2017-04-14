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
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices.ComTypes;
using System.Security.AccessControl;
using System.Security.Principal;

namespace AutomatedStudentRecordKeeper
{
    public partial class AddCourse : Form
    {
        //Variable Declaration
        string selectedInputFile;
        NpgsqlCommand cmd = new NpgsqlCommand();
        string courseSection = string.Empty;
        string yearSection = string.Empty;
        int yearLevel = 0;
        int firstUsedYear = 0;

        public AddCourse()
        {
            InitializeComponent();
            //Populate dropdownbox with years
            int tempyear = DateTime.Now.Year;
            for (int i = 6; i >= -1; i--)
            {
                this.yeardropbox.Items.Add((tempyear - i).ToString() + "/" + (tempyear - i + 1).ToString());
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            AddCourseTable.Hide();
            //connection string
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (conn.State == System.Data.ConnectionState.Open)
            {
                //checks for valid information
                if (Yearleveldropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Year level");

                }
                else if (sectiondropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a section");
                }
                else if (yeardropbox.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Year");
                }
                else
                {
                    int count = 0;
                    for (int j = 0; j < this.AddCourseTable.RowCount; j++)
                    {
                        if (string.IsNullOrWhiteSpace(AddCourseTable.GetControlFromPosition(0, j).Text) ||
                            string.IsNullOrWhiteSpace(AddCourseTable.GetControlFromPosition(1, j).Text) ||
                            string.IsNullOrWhiteSpace(AddCourseTable.GetControlFromPosition(2, j).Text) ||
                            string.IsNullOrWhiteSpace(AddCourseTable.GetControlFromPosition(3, j).Text))
                        {

                        }
                        else if (AddCourseTable.GetControlFromPosition(0, j).Text.Length != 4 || AddCourseTable.GetControlFromPosition(1, j).Text.Length != 4)
                        {

                        }
                        else
                        {
                            NpgsqlDataReader reader;
                            NpgsqlCommand cmd;
                            string checkifexists = "False";
                            //checks if data already exists
                            cmd = new NpgsqlCommand("select exists(select true from courses where coursesubject = :sub and coursenumber = :num and yearsection = :year)", conn);
                            cmd.Parameters.Add(new NpgsqlParameter("sub", AddCourseTable.GetControlFromPosition(0, j).Text));
                            cmd.Parameters.Add(new NpgsqlParameter("num", AddCourseTable.GetControlFromPosition(1, j).Text));
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
                                //query for course entry into database
                                cmd = new NpgsqlCommand("insert into courses values(:sub, :num, :sec, :name, :cred, :yrlvl, :yrsec, :entyear, 'curric')", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("sub", AddCourseTable.GetControlFromPosition(0, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("num", AddCourseTable.GetControlFromPosition(1, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("sec", sectiondropbox.Text));
                                cmd.Parameters.Add(new NpgsqlParameter("name", AddCourseTable.GetControlFromPosition(2, j).Text));
                                cmd.Parameters.Add(new NpgsqlParameter("cred", double.Parse(AddCourseTable.GetControlFromPosition(3, j).Text)));
                                cmd.Parameters.Add(new NpgsqlParameter("yrlvl", int.Parse(Yearleveldropbox.Text)));
                                cmd.Parameters.Add(new NpgsqlParameter("yrsec", yeardropbox.Text));
                                cmd.Parameters.Add(new NpgsqlParameter("entyear", int.Parse(yeardropbox.Text.Substring(0, 4))));
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (NpgsqlException ex)
                                {

                                }
                            }
                            //clear table after complete query
                            AddCourseTable.GetControlFromPosition(0, j).Text = "";
                            AddCourseTable.GetControlFromPosition(1, j).Text = "";
                            AddCourseTable.GetControlFromPosition(2, j).Text = "";
                            AddCourseTable.GetControlFromPosition(3, j).Text = "";
                            count++;
                        }
                    }
                    conn.Close();
                    //Show message after complete
                    MessageBox.Show(count.ToString() + " rows added to table, check formating if form not cleared");
                }
            }
            else
            {
                MessageBox.Show("Connection error to database");
            }
            AddCourseTable.Show();
        }
        //key press to only allow uppercase letters
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
        //key press to only allow numbers
        private void richTextBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }

        private void import_courses_Click(object sender, EventArgs e)
        {
            NpgsqlConnection conn = new NpgsqlConnection("Server=Localhost; Port=5432; Database=studentrecordkeeper; User Id=postgres; Password=;");
            //connect to database
            conn.Open();
            if (yeardropbox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Year");
            }
            else
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    firstUsedYear = int.Parse(yeardropbox.Text.Substring(0, 4));
                    OpenFileDialog choofdlog = new OpenFileDialog(); //opens file viewer
                    choofdlog.Filter = "doc files (*.docx*)|*.docx|(*.doc*)|*.doc"; //only shows doc files
                    choofdlog.Title = "Select a Document File";
                    choofdlog.FilterIndex = 1;
                    choofdlog.Multiselect = false; //one file at a time

                    if (choofdlog.ShowDialog() == DialogResult.OK)
                    {
                        selectedInputFile = choofdlog.FileName; //sets path

                        Debug.Write("START");
                        //initializes and shows wait screen
                        waitScreen waitscrn = new waitScreen();
                        waitscrn.Show();
                        System.Windows.Forms.Application.DoEvents();

                        //loads document file using microsoft word without showing microsoft word
                        List<string> data = new List<string>();
                        Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                        Document document = word.Documents.Open(selectedInputFile, ReadOnly: true);

                        foreach (Paragraph objParagraph in document.Paragraphs)
                            data.Add(objParagraph.Range.Text.Trim());

                        string totaltext = string.Join(string.Empty, data.ToArray());

                        //closes document and exits word
                        ((_Document)document).Close();
                        ((_Application)word).Quit();

                        //cleaning of file
                        totaltext = totaltext.Substring(totaltext.IndexOf("YEAR")); //removes irrelevant info at beginning of doc file
                        totaltext = Regex.Replace(totaltext, @"\a", string.Empty);
                        totaltext = Regex.Replace(totaltext, @"(Lab)\r|(Lec)\r", string.Empty);
                        totaltext = Regex.Replace(totaltext, @"\d\.\d", string.Empty);
                        totaltext = Regex.Replace(totaltext, @"\b\d{1}\s|\b\d{2}\s", Environment.NewLine);
                        totaltext = Regex.Replace(totaltext, @"YEAR|FALL |WINTER ", string.Empty, RegexOptions.Singleline);
                        totaltext = Regex.Replace(totaltext, @"SECOND|THIRD|FOURTH", string.Empty);
                        totaltext = Regex.Replace(totaltext, @"(Total).*?(Hours)", string.Empty);
                        totaltext = Regex.Replace(totaltext, @"(Sociology 2755).*?\n", string.Empty);
                        totaltext = Regex.Replace(totaltext, Environment.NewLine, string.Empty);
                        totaltext = Regex.Replace(totaltext, @"\r+", "\n");
                        totaltext = Regex.Replace(totaltext, @"(?<=([\d]{4}))\s+", "~", RegexOptions.Singleline);
                        totaltext = Regex.Replace(totaltext, @"(One half course from Science Elective Course List).*?\n", string.Empty, RegexOptions.Singleline);
                        totaltext = Regex.Replace(totaltext, @"(One half course from Engineering Elective Course List).*?\n", string.Empty, RegexOptions.Singleline);
                        totaltext = Regex.Replace(totaltext, @"(One complementary).*?\n", string.Empty, RegexOptions.Singleline);

                        string[] term = totaltext.Split(new string[] { "TERM " }, StringSplitOptions.RemoveEmptyEntries);

                        string coursesFile = System.IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "courses.txt");

                        //deletes file if exists and recreates it
                        System.IO.File.Delete(coursesFile);
                        using (System.IO.File.Create(coursesFile)) { }

                        //sets permissions for created file
                        FileSecurity access = System.IO.File.GetAccessControl(coursesFile);
                        SecurityIdentifier everyone = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
                        access.AddAccessRule(new FileSystemAccessRule(everyone, FileSystemRights.ReadAndExecute, AccessControlType.Allow));
                        System.IO.File.SetAccessControl(coursesFile, access);

                        for (int j = 0; j < term.Length; j++)
                        {
                            term[j] = Regex.Replace(term[j], @"(?<=([A-Z]{4}))\s+", "~");
                            term[j] = Regex.Replace(term[j], @"(- ).*?(term\n)", string.Empty);

                            var courseNumber = Regex.Matches(term[j], @"[A-Z]{4}.*?\n", RegexOptions.Singleline);
                            var courseNumberList = courseNumber.Cast<Match>().Select(match => match.Value).ToList();

                            //look into getting rid of the empty first file to fix this and have j % 2 == 0 instead
                            //may save from future errors if something in the original file changes later on?
                            if (j % 2 == 1)
                            {
                                courseSection = "~Fall";
                                yearLevel++;
                            }
                            else
                                courseSection = "~Winter";

                            yearSection = (firstUsedYear.ToString()) + "/" + ((firstUsedYear + 1).ToString());

                            for (int i = 0; i < courseNumberList.Count; i++)
                            {
                                courseNumberList[i] += courseSection + "~0.5~" + yearLevel + "~" + yearSection + "~" + firstUsedYear + "~curric";
                                courseNumberList[i] = Regex.Replace(courseNumberList[i], @"[\n\r]+", string.Empty);
                            }

                            System.IO.File.AppendAllLines(coursesFile, courseNumberList);
                            // Add the access control entry to the file.

                        }

                        try
                        {
                            cmd = new NpgsqlCommand("COPY courses (coursesubject, coursenumber, coursename, coursesection,"
                            + "credits, yearlevel, yearsection, yearused, type ) FROM "
                            + "'" + coursesFile + "'" + " DELIMITER '~' CSV", conn);
                            cmd.ExecuteNonQuery();
                        }
                        catch
                        {
                            goto ERROR;
                        }

                        waitscrn.Hide();
                        this.Activate();
                        var result = MessageBox.Show("Importing of software engineering course load to the database was successful. Would you like to import another file?", "SUCCESS",
                             MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                            this.Activate();
                        if (result == DialogResult.No)
                            this.Close();
                        goto END;
                    ERROR:
                        waitscrn.Hide();
                        this.Activate();
                        MessageBox.Show("Error Adding Courses to DataBase", "ERROR");
                        this.Close();
                    END:
                        Debug.WriteLine("DONE");
                    }
                    else
                        selectedInputFile = string.Empty;
                }
            }
        }
    }
}
