namespace AutomatedStudentRecordKeeper
{
    partial class AddStudent
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.AddStudentButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Namebox = new System.Windows.Forms.TextBox();
            this.StudentNumberBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.PrevSchoolBox = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.PrevProgramBox = new System.Windows.Forms.TextBox();
            this.Yeardropbox = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.MakeupCourseTable = new System.Windows.Forms.TableLayoutPanel();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.richTextBox4 = new System.Windows.Forms.RichTextBox();
            this.richTextBox5 = new System.Windows.Forms.RichTextBox();
            this.richTextBox6 = new System.Windows.Forms.RichTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.AddRowMakeup = new System.Windows.Forms.Button();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.AddRowCredits = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.CourseCredTable = new System.Windows.Forms.TableLayoutPanel();
            this.richTextBox7 = new System.Windows.Forms.RichTextBox();
            this.richTextBox8 = new System.Windows.Forms.RichTextBox();
            this.richTextBox9 = new System.Windows.Forms.RichTextBox();
            this.richTextBox10 = new System.Windows.Forms.RichTextBox();
            this.richTextBox11 = new System.Windows.Forms.RichTextBox();
            this.richTextBox12 = new System.Windows.Forms.RichTextBox();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel6 = new System.Windows.Forms.TableLayoutPanel();
            this.MakeupCourseTable.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.CourseCredTable.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel6.SuspendLayout();
            this.SuspendLayout();
            // 
            // AddStudentButton
            // 
            this.AddStudentButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.AddStudentButton.BackColor = System.Drawing.SystemColors.Window;
            this.AddStudentButton.Location = new System.Drawing.Point(350, 202);
            this.AddStudentButton.Name = "AddStudentButton";
            this.AddStudentButton.Size = new System.Drawing.Size(75, 23);
            this.AddStudentButton.TabIndex = 0;
            this.AddStudentButton.Text = "Confirm";
            this.AddStudentButton.UseVisualStyleBackColor = false;
            this.AddStudentButton.Click += new System.EventHandler(this.AddStudentButton_Click);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(188, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Name";
            // 
            // Namebox
            // 
            this.Namebox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.Namebox.Location = new System.Drawing.Point(229, 37);
            this.Namebox.Name = "Namebox";
            this.Namebox.Size = new System.Drawing.Size(133, 20);
            this.Namebox.TabIndex = 2;
            this.Namebox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Namebox_KeyPress);
            // 
            // StudentNumberBox
            // 
            this.StudentNumberBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.StudentNumberBox.Location = new System.Drawing.Point(458, 37);
            this.StudentNumberBox.MaxLength = 7;
            this.StudentNumberBox.Name = "StudentNumberBox";
            this.StudentNumberBox.Size = new System.Drawing.Size(130, 20);
            this.StudentNumberBox.TabIndex = 3;
            this.StudentNumberBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(368, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Student Number";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(172, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 26);
            this.label3.TabIndex = 5;
            this.label3.Text = "Previous \r\nSchool";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // PrevSchoolBox
            // 
            this.PrevSchoolBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.PrevSchoolBox.Location = new System.Drawing.Point(229, 71);
            this.PrevSchoolBox.Name = "PrevSchoolBox";
            this.PrevSchoolBox.Size = new System.Drawing.Size(133, 20);
            this.PrevSchoolBox.TabIndex = 6;
            this.PrevSchoolBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Namebox_KeyPress);
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(368, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Previous Program";
            // 
            // PrevProgramBox
            // 
            this.PrevProgramBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.PrevProgramBox.Location = new System.Drawing.Point(458, 71);
            this.PrevProgramBox.Name = "PrevProgramBox";
            this.PrevProgramBox.Size = new System.Drawing.Size(130, 20);
            this.PrevProgramBox.TabIndex = 8;
            this.PrevProgramBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Namebox_KeyPress);
            // 
            // Yeardropbox
            // 
            this.Yeardropbox.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.Yeardropbox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Yeardropbox.FormattingEnabled = true;
            this.Yeardropbox.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4"});
            this.Yeardropbox.Location = new System.Drawing.Point(402, 101);
            this.Yeardropbox.Name = "Yeardropbox";
            this.Yeardropbox.Size = new System.Drawing.Size(54, 21);
            this.Yeardropbox.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(338, 104);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(58, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Year Level";
            // 
            // MakeupCourseTable
            // 
            this.MakeupCourseTable.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.MakeupCourseTable.AutoScroll = true;
            this.MakeupCourseTable.AutoSize = true;
            this.MakeupCourseTable.BackColor = System.Drawing.SystemColors.Window;
            this.MakeupCourseTable.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.MakeupCourseTable.ColumnCount = 3;
            this.MakeupCourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 46.15385F));
            this.MakeupCourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 53.84615F));
            this.MakeupCourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 244F));
            this.MakeupCourseTable.Controls.Add(this.richTextBox1, 0, 0);
            this.MakeupCourseTable.Controls.Add(this.richTextBox2, 1, 0);
            this.MakeupCourseTable.Controls.Add(this.richTextBox3, 2, 0);
            this.MakeupCourseTable.Controls.Add(this.richTextBox4, 0, 1);
            this.MakeupCourseTable.Controls.Add(this.richTextBox5, 1, 1);
            this.MakeupCourseTable.Controls.Add(this.richTextBox6, 2, 1);
            this.MakeupCourseTable.Location = new System.Drawing.Point(3, 23);
            this.MakeupCourseTable.Name = "MakeupCourseTable";
            this.MakeupCourseTable.RowCount = 2;
            this.MakeupCourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.90909F));
            this.MakeupCourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.09091F));
            this.MakeupCourseTable.Size = new System.Drawing.Size(359, 58);
            this.MakeupCourseTable.TabIndex = 11;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(4, 4);
            this.richTextBox1.MaxLength = 4;
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(45, 21);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            this.richTextBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(56, 4);
            this.richTextBox2.MaxLength = 4;
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(53, 21);
            this.richTextBox2.TabIndex = 1;
            this.richTextBox2.Text = "";
            this.richTextBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(116, 4);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(233, 21);
            this.richTextBox3.TabIndex = 2;
            this.richTextBox3.Text = "";
            // 
            // richTextBox4
            // 
            this.richTextBox4.Location = new System.Drawing.Point(4, 32);
            this.richTextBox4.MaxLength = 4;
            this.richTextBox4.Name = "richTextBox4";
            this.richTextBox4.Size = new System.Drawing.Size(45, 21);
            this.richTextBox4.TabIndex = 3;
            this.richTextBox4.Text = "";
            this.richTextBox4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            // 
            // richTextBox5
            // 
            this.richTextBox5.Location = new System.Drawing.Point(56, 32);
            this.richTextBox5.MaxLength = 4;
            this.richTextBox5.Name = "richTextBox5";
            this.richTextBox5.Size = new System.Drawing.Size(53, 21);
            this.richTextBox5.TabIndex = 4;
            this.richTextBox5.Text = "";
            this.richTextBox5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            // 
            // richTextBox6
            // 
            this.richTextBox6.Location = new System.Drawing.Point(116, 32);
            this.richTextBox6.Name = "richTextBox6";
            this.richTextBox6.Size = new System.Drawing.Size(233, 21);
            this.richTextBox6.TabIndex = 5;
            this.richTextBox6.Text = "";
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(139, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Make Up Courses";
            // 
            // label7
            // 
            this.label7.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Subject";
            // 
            // label8
            // 
            this.label8.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(201, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 13);
            this.label8.TabIndex = 14;
            this.label8.Text = "Course Title";
            // 
            // label11
            // 
            this.label11.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(519, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(75, 13);
            this.label11.TabIndex = 16;
            this.label11.Text = "Course Credits";
            // 
            // label12
            // 
            this.label12.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(63, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(44, 13);
            this.label12.TabIndex = 23;
            this.label12.Text = "Number";
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.AutoSize = true;
            this.tableLayoutPanel4.ColumnCount = 1;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Controls.Add(this.AddRowMakeup, 0, 2);
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel5, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.MakeupCourseTable, 0, 1);
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 16);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 3;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(365, 121);
            this.tableLayoutPanel4.TabIndex = 30;
            // 
            // AddRowMakeup
            // 
            this.AddRowMakeup.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.AddRowMakeup.BackColor = System.Drawing.SystemColors.Window;
            this.AddRowMakeup.Location = new System.Drawing.Point(142, 87);
            this.AddRowMakeup.Name = "AddRowMakeup";
            this.AddRowMakeup.Size = new System.Drawing.Size(81, 23);
            this.AddRowMakeup.TabIndex = 32;
            this.AddRowMakeup.Text = "Add Row";
            this.AddRowMakeup.UseVisualStyleBackColor = false;
            this.AddRowMakeup.Click += new System.EventHandler(this.AddRowMakeup_Click);
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.ColumnCount = 3;
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.30508F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.69492F));
            this.tableLayoutPanel5.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 234F));
            this.tableLayoutPanel5.Controls.Add(this.label12, 1, 0);
            this.tableLayoutPanel5.Controls.Add(this.label7, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.label8, 2, 0);
            this.tableLayoutPanel5.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.RowCount = 1;
            this.tableLayoutPanel5.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel5.Size = new System.Drawing.Size(350, 14);
            this.tableLayoutPanel5.TabIndex = 31;
            // 
            // AddRowCredits
            // 
            this.AddRowCredits.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.AddRowCredits.BackColor = System.Drawing.SystemColors.Window;
            this.AddRowCredits.Location = new System.Drawing.Point(142, 87);
            this.AddRowCredits.Name = "AddRowCredits";
            this.AddRowCredits.Size = new System.Drawing.Size(81, 23);
            this.AddRowCredits.TabIndex = 32;
            this.AddRowCredits.Text = "Add Row";
            this.AddRowCredits.UseVisualStyleBackColor = false;
            this.AddRowCredits.Click += new System.EventHandler(this.AddRowCredits_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.AddRowCredits, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.CourseCredTable, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(374, 16);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 37F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(365, 121);
            this.tableLayoutPanel1.TabIndex = 33;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.30508F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.69492F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 234F));
            this.tableLayoutPanel2.Controls.Add(this.label9, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label10, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label13, 2, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(350, 14);
            this.tableLayoutPanel2.TabIndex = 31;
            // 
            // label9
            // 
            this.label9.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(63, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 13);
            this.label9.TabIndex = 23;
            this.label9.Text = "Number";
            // 
            // label10
            // 
            this.label10.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 13;
            this.label10.Text = "Subject";
            // 
            // label13
            // 
            this.label13.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(201, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(63, 13);
            this.label13.TabIndex = 14;
            this.label13.Text = "Course Title";
            // 
            // CourseCredTable
            // 
            this.CourseCredTable.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.CourseCredTable.AutoScroll = true;
            this.CourseCredTable.AutoSize = true;
            this.CourseCredTable.BackColor = System.Drawing.SystemColors.Window;
            this.CourseCredTable.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.CourseCredTable.ColumnCount = 3;
            this.CourseCredTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 46.15385F));
            this.CourseCredTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 53.84615F));
            this.CourseCredTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 244F));
            this.CourseCredTable.Controls.Add(this.richTextBox7, 0, 0);
            this.CourseCredTable.Controls.Add(this.richTextBox8, 1, 0);
            this.CourseCredTable.Controls.Add(this.richTextBox9, 2, 0);
            this.CourseCredTable.Controls.Add(this.richTextBox10, 0, 1);
            this.CourseCredTable.Controls.Add(this.richTextBox11, 1, 1);
            this.CourseCredTable.Controls.Add(this.richTextBox12, 2, 1);
            this.CourseCredTable.Location = new System.Drawing.Point(3, 23);
            this.CourseCredTable.Name = "CourseCredTable";
            this.CourseCredTable.RowCount = 2;
            this.CourseCredTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.90909F));
            this.CourseCredTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.09091F));
            this.CourseCredTable.Size = new System.Drawing.Size(359, 58);
            this.CourseCredTable.TabIndex = 11;
            // 
            // richTextBox7
            // 
            this.richTextBox7.Location = new System.Drawing.Point(4, 4);
            this.richTextBox7.MaxLength = 4;
            this.richTextBox7.Name = "richTextBox7";
            this.richTextBox7.Size = new System.Drawing.Size(45, 21);
            this.richTextBox7.TabIndex = 0;
            this.richTextBox7.Text = "";
            this.richTextBox7.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            // 
            // richTextBox8
            // 
            this.richTextBox8.Location = new System.Drawing.Point(56, 4);
            this.richTextBox8.MaxLength = 4;
            this.richTextBox8.Name = "richTextBox8";
            this.richTextBox8.Size = new System.Drawing.Size(53, 21);
            this.richTextBox8.TabIndex = 1;
            this.richTextBox8.Text = "";
            this.richTextBox8.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            // 
            // richTextBox9
            // 
            this.richTextBox9.Location = new System.Drawing.Point(116, 4);
            this.richTextBox9.Name = "richTextBox9";
            this.richTextBox9.Size = new System.Drawing.Size(233, 21);
            this.richTextBox9.TabIndex = 2;
            this.richTextBox9.Text = "";
            // 
            // richTextBox10
            // 
            this.richTextBox10.Location = new System.Drawing.Point(4, 32);
            this.richTextBox10.MaxLength = 4;
            this.richTextBox10.Name = "richTextBox10";
            this.richTextBox10.Size = new System.Drawing.Size(45, 21);
            this.richTextBox10.TabIndex = 3;
            this.richTextBox10.Text = "";
            this.richTextBox10.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox1_KeyPress);
            // 
            // richTextBox11
            // 
            this.richTextBox11.Location = new System.Drawing.Point(56, 32);
            this.richTextBox11.MaxLength = 4;
            this.richTextBox11.Name = "richTextBox11";
            this.richTextBox11.Size = new System.Drawing.Size(53, 21);
            this.richTextBox11.TabIndex = 4;
            this.richTextBox11.Text = "";
            this.richTextBox11.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.StudentNumberBox_KeyPress);
            // 
            // richTextBox12
            // 
            this.richTextBox12.Location = new System.Drawing.Point(116, 32);
            this.richTextBox12.Name = "richTextBox12";
            this.richTextBox12.Size = new System.Drawing.Size(233, 21);
            this.richTextBox12.TabIndex = 5;
            this.richTextBox12.Text = "";
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.tableLayoutPanel3.AutoSize = true;
            this.tableLayoutPanel3.ColumnCount = 2;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel3.Controls.Add(this.label6, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.label11, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel1, 1, 1);
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel4, 0, 1);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(17, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 6.779661F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 93.22034F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(742, 192);
            this.tableLayoutPanel3.TabIndex = 34;
            // 
            // tableLayoutPanel6
            // 
            this.tableLayoutPanel6.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.tableLayoutPanel6.AutoSize = true;
            this.tableLayoutPanel6.ColumnCount = 1;
            this.tableLayoutPanel6.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel6.Controls.Add(this.tableLayoutPanel3, 0, 0);
            this.tableLayoutPanel6.Controls.Add(this.AddStudentButton, 0, 1);
            this.tableLayoutPanel6.Location = new System.Drawing.Point(12, 128);
            this.tableLayoutPanel6.Name = "tableLayoutPanel6";
            this.tableLayoutPanel6.RowCount = 2;
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 86.93877F));
            this.tableLayoutPanel6.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 13.06122F));
            this.tableLayoutPanel6.Size = new System.Drawing.Size(776, 228);
            this.tableLayoutPanel6.TabIndex = 35;
            this.tableLayoutPanel6.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel6_Paint);
            // 
            // AddStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 397);
            this.Controls.Add(this.tableLayoutPanel6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Yeardropbox);
            this.Controls.Add(this.PrevProgramBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.PrevSchoolBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.StudentNumberBox);
            this.Controls.Add(this.Namebox);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "AddStudent";
            this.Text = "TransferStudent";
            this.Load += new System.EventHandler(this.TransferStudent_Load);
            this.MakeupCourseTable.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel5.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.CourseCredTable.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            this.tableLayoutPanel6.ResumeLayout(false);
            this.tableLayoutPanel6.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button AddStudentButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Namebox;
        private System.Windows.Forms.TextBox StudentNumberBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox PrevSchoolBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox PrevProgramBox;
        private System.Windows.Forms.ComboBox Yeardropbox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TableLayoutPanel MakeupCourseTable;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.RichTextBox richTextBox4;
        private System.Windows.Forms.RichTextBox richTextBox5;
        private System.Windows.Forms.RichTextBox richTextBox6;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.Button AddRowMakeup;
        private System.Windows.Forms.Button AddRowCredits;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TableLayoutPanel CourseCredTable;
        private System.Windows.Forms.RichTextBox richTextBox7;
        private System.Windows.Forms.RichTextBox richTextBox8;
        private System.Windows.Forms.RichTextBox richTextBox9;
        private System.Windows.Forms.RichTextBox richTextBox10;
        private System.Windows.Forms.RichTextBox richTextBox11;
        private System.Windows.Forms.RichTextBox richTextBox12;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel6;
    }
}