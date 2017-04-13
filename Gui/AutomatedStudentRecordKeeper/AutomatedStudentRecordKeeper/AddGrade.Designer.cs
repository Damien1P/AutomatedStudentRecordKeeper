using System;
namespace AutomatedStudentRecordKeeper
{
    partial class AddGrade
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
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.CourseTable = new System.Windows.Forms.TableLayoutPanel();
            this.richTextBox28 = new System.Windows.Forms.RichTextBox();
            this.richTextBox27 = new System.Windows.Forms.RichTextBox();
            this.richTextBox26 = new System.Windows.Forms.RichTextBox();
            this.richTextBox24 = new System.Windows.Forms.RichTextBox();
            this.richTextBox23 = new System.Windows.Forms.RichTextBox();
            this.richTextBox12 = new System.Windows.Forms.RichTextBox();
            this.richTextBox22 = new System.Windows.Forms.RichTextBox();
            this.richTextBox20 = new System.Windows.Forms.RichTextBox();
            this.richTextBox19 = new System.Windows.Forms.RichTextBox();
            this.richTextBox18 = new System.Windows.Forms.RichTextBox();
            this.richTextBox16 = new System.Windows.Forms.RichTextBox();
            this.richTextBox15 = new System.Windows.Forms.RichTextBox();
            this.richTextBox14 = new System.Windows.Forms.RichTextBox();
            this.richTextBox11 = new System.Windows.Forms.RichTextBox();
            this.richTextBox10 = new System.Windows.Forms.RichTextBox();
            this.richTextBox8 = new System.Windows.Forms.RichTextBox();
            this.richTextBox7 = new System.Windows.Forms.RichTextBox();
            this.richTextBox6 = new System.Windows.Forms.RichTextBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.richTextBox4 = new System.Windows.Forms.RichTextBox();
            this.StudentNumber = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.yeardropbox = new System.Windows.Forms.ComboBox();
            this.importHTML = new System.Windows.Forms.Button();
            this.CourseTable.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Course Subject";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(293, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(36, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Grade";
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button1.Location = new System.Drawing.Point(74, 352);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "Submit";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(157, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Course Number";
            // 
            // CourseTable
            // 
            this.CourseTable.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.CourseTable.BackColor = System.Drawing.SystemColors.Window;
            this.CourseTable.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Inset;
            this.CourseTable.ColumnCount = 3;
            this.CourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.85185F));
            this.CourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.CourseTable.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 96F));
            this.CourseTable.Controls.Add(this.richTextBox28, 2, 6);
            this.CourseTable.Controls.Add(this.richTextBox27, 1, 6);
            this.CourseTable.Controls.Add(this.richTextBox26, 0, 6);
            this.CourseTable.Controls.Add(this.richTextBox24, 2, 5);
            this.CourseTable.Controls.Add(this.richTextBox23, 1, 5);
            this.CourseTable.Controls.Add(this.richTextBox12, 2, 2);
            this.CourseTable.Controls.Add(this.richTextBox22, 0, 5);
            this.CourseTable.Controls.Add(this.richTextBox20, 2, 4);
            this.CourseTable.Controls.Add(this.richTextBox19, 1, 4);
            this.CourseTable.Controls.Add(this.richTextBox18, 0, 4);
            this.CourseTable.Controls.Add(this.richTextBox16, 2, 3);
            this.CourseTable.Controls.Add(this.richTextBox15, 1, 3);
            this.CourseTable.Controls.Add(this.richTextBox14, 0, 3);
            this.CourseTable.Controls.Add(this.richTextBox11, 1, 2);
            this.CourseTable.Controls.Add(this.richTextBox10, 0, 2);
            this.CourseTable.Controls.Add(this.richTextBox8, 2, 1);
            this.CourseTable.Controls.Add(this.richTextBox7, 1, 1);
            this.CourseTable.Controls.Add(this.richTextBox6, 0, 1);
            this.CourseTable.Controls.Add(this.richTextBox2, 0, 0);
            this.CourseTable.Controls.Add(this.richTextBox3, 1, 0);
            this.CourseTable.Controls.Add(this.richTextBox4, 2, 0);
            this.CourseTable.Location = new System.Drawing.Point(12, 100);
            this.CourseTable.Name = "CourseTable";
            this.CourseTable.RowCount = 7;
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.CourseTable.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CourseTable.Size = new System.Drawing.Size(335, 245);
            this.CourseTable.TabIndex = 18;
            // 
            // richTextBox28
            // 
            this.richTextBox28.Location = new System.Drawing.Point(240, 209);
            this.richTextBox28.MaxLength = 3;
            this.richTextBox28.Name = "richTextBox28";
            this.richTextBox28.Size = new System.Drawing.Size(80, 26);
            this.richTextBox28.TabIndex = 27;
            this.richTextBox28.Text = "";
            this.richTextBox28.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox27
            // 
            this.richTextBox27.Location = new System.Drawing.Point(118, 209);
            this.richTextBox27.MaxLength = 4;
            this.richTextBox27.Name = "richTextBox27";
            this.richTextBox27.Size = new System.Drawing.Size(114, 26);
            this.richTextBox27.TabIndex = 26;
            this.richTextBox27.Text = "";
            this.richTextBox27.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox26
            // 
            this.richTextBox26.Location = new System.Drawing.Point(5, 209);
            this.richTextBox26.MaxLength = 4;
            this.richTextBox26.Name = "richTextBox26";
            this.richTextBox26.Size = new System.Drawing.Size(105, 26);
            this.richTextBox26.TabIndex = 25;
            this.richTextBox26.Text = "";
            this.richTextBox26.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox24
            // 
            this.richTextBox24.Location = new System.Drawing.Point(240, 175);
            this.richTextBox24.MaxLength = 3;
            this.richTextBox24.Name = "richTextBox24";
            this.richTextBox24.Size = new System.Drawing.Size(80, 26);
            this.richTextBox24.TabIndex = 23;
            this.richTextBox24.Text = "";
            this.richTextBox24.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox23
            // 
            this.richTextBox23.Location = new System.Drawing.Point(118, 175);
            this.richTextBox23.MaxLength = 4;
            this.richTextBox23.Name = "richTextBox23";
            this.richTextBox23.Size = new System.Drawing.Size(114, 26);
            this.richTextBox23.TabIndex = 22;
            this.richTextBox23.Text = "";
            this.richTextBox23.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox12
            // 
            this.richTextBox12.Location = new System.Drawing.Point(240, 73);
            this.richTextBox12.MaxLength = 3;
            this.richTextBox12.Name = "richTextBox12";
            this.richTextBox12.Size = new System.Drawing.Size(80, 26);
            this.richTextBox12.TabIndex = 11;
            this.richTextBox12.Text = "";
            this.richTextBox12.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox22
            // 
            this.richTextBox22.Location = new System.Drawing.Point(5, 175);
            this.richTextBox22.MaxLength = 4;
            this.richTextBox22.Name = "richTextBox22";
            this.richTextBox22.Size = new System.Drawing.Size(105, 26);
            this.richTextBox22.TabIndex = 21;
            this.richTextBox22.Text = "";
            this.richTextBox22.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox20
            // 
            this.richTextBox20.Location = new System.Drawing.Point(240, 141);
            this.richTextBox20.MaxLength = 3;
            this.richTextBox20.Name = "richTextBox20";
            this.richTextBox20.Size = new System.Drawing.Size(80, 26);
            this.richTextBox20.TabIndex = 19;
            this.richTextBox20.Text = "";
            this.richTextBox20.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox19
            // 
            this.richTextBox19.Location = new System.Drawing.Point(118, 141);
            this.richTextBox19.MaxLength = 4;
            this.richTextBox19.Name = "richTextBox19";
            this.richTextBox19.Size = new System.Drawing.Size(114, 26);
            this.richTextBox19.TabIndex = 18;
            this.richTextBox19.Text = "";
            this.richTextBox19.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox18
            // 
            this.richTextBox18.Location = new System.Drawing.Point(5, 141);
            this.richTextBox18.MaxLength = 4;
            this.richTextBox18.Name = "richTextBox18";
            this.richTextBox18.Size = new System.Drawing.Size(105, 26);
            this.richTextBox18.TabIndex = 17;
            this.richTextBox18.Text = "";
            this.richTextBox18.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox16
            // 
            this.richTextBox16.Location = new System.Drawing.Point(240, 107);
            this.richTextBox16.MaxLength = 3;
            this.richTextBox16.Name = "richTextBox16";
            this.richTextBox16.Size = new System.Drawing.Size(80, 26);
            this.richTextBox16.TabIndex = 15;
            this.richTextBox16.Text = "";
            this.richTextBox16.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox15
            // 
            this.richTextBox15.Location = new System.Drawing.Point(118, 107);
            this.richTextBox15.MaxLength = 4;
            this.richTextBox15.Name = "richTextBox15";
            this.richTextBox15.Size = new System.Drawing.Size(114, 26);
            this.richTextBox15.TabIndex = 14;
            this.richTextBox15.Text = "";
            this.richTextBox15.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox14
            // 
            this.richTextBox14.Location = new System.Drawing.Point(5, 107);
            this.richTextBox14.MaxLength = 4;
            this.richTextBox14.Name = "richTextBox14";
            this.richTextBox14.Size = new System.Drawing.Size(105, 26);
            this.richTextBox14.TabIndex = 13;
            this.richTextBox14.Text = "";
            this.richTextBox14.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox11
            // 
            this.richTextBox11.Location = new System.Drawing.Point(118, 73);
            this.richTextBox11.MaxLength = 4;
            this.richTextBox11.Name = "richTextBox11";
            this.richTextBox11.Size = new System.Drawing.Size(114, 26);
            this.richTextBox11.TabIndex = 10;
            this.richTextBox11.Text = "";
            this.richTextBox11.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox10
            // 
            this.richTextBox10.Location = new System.Drawing.Point(5, 73);
            this.richTextBox10.MaxLength = 4;
            this.richTextBox10.Name = "richTextBox10";
            this.richTextBox10.Size = new System.Drawing.Size(105, 26);
            this.richTextBox10.TabIndex = 9;
            this.richTextBox10.Text = "";
            this.richTextBox10.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox8
            // 
            this.richTextBox8.Location = new System.Drawing.Point(240, 39);
            this.richTextBox8.MaxLength = 3;
            this.richTextBox8.Name = "richTextBox8";
            this.richTextBox8.Size = new System.Drawing.Size(80, 26);
            this.richTextBox8.TabIndex = 7;
            this.richTextBox8.Text = "";
            this.richTextBox8.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox7
            // 
            this.richTextBox7.Location = new System.Drawing.Point(118, 39);
            this.richTextBox7.MaxLength = 4;
            this.richTextBox7.Name = "richTextBox7";
            this.richTextBox7.Size = new System.Drawing.Size(114, 26);
            this.richTextBox7.TabIndex = 6;
            this.richTextBox7.Text = "";
            this.richTextBox7.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox6
            // 
            this.richTextBox6.Location = new System.Drawing.Point(5, 39);
            this.richTextBox6.MaxLength = 4;
            this.richTextBox6.Name = "richTextBox6";
            this.richTextBox6.Size = new System.Drawing.Size(105, 26);
            this.richTextBox6.TabIndex = 5;
            this.richTextBox6.Text = "";
            this.richTextBox6.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox2
            // 
            this.richTextBox2.Location = new System.Drawing.Point(5, 5);
            this.richTextBox2.MaxLength = 4;
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(105, 26);
            this.richTextBox2.TabIndex = 1;
            this.richTextBox2.Text = "";
            this.richTextBox2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox2_KeyPress);
            // 
            // richTextBox3
            // 
            this.richTextBox3.Location = new System.Drawing.Point(118, 5);
            this.richTextBox3.MaxLength = 4;
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(114, 26);
            this.richTextBox3.TabIndex = 2;
            this.richTextBox3.Text = "";
            this.richTextBox3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // richTextBox4
            // 
            this.richTextBox4.Location = new System.Drawing.Point(240, 5);
            this.richTextBox4.MaxLength = 3;
            this.richTextBox4.Name = "richTextBox4";
            this.richTextBox4.Size = new System.Drawing.Size(80, 26);
            this.richTextBox4.TabIndex = 3;
            this.richTextBox4.Text = "";
            this.richTextBox4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // StudentNumber
            // 
            this.StudentNumber.Location = new System.Drawing.Point(140, 12);
            this.StudentNumber.MaxLength = 7;
            this.StudentNumber.Name = "StudentNumber";
            this.StudentNumber.Size = new System.Drawing.Size(139, 20);
            this.StudentNumber.TabIndex = 19;
            this.StudentNumber.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.richTextBox3_KeyPress);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(52, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 20;
            this.label2.Text = "Student Number";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(148, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Year";
            // 
            // yeardropbox
            // 
            this.yeardropbox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.yeardropbox.FormattingEnabled = true;
            this.yeardropbox.Location = new System.Drawing.Point(183, 47);
            this.yeardropbox.Name = "yeardropbox";
            this.yeardropbox.Size = new System.Drawing.Size(54, 21);
            this.yeardropbox.TabIndex = 21;
            this.yeardropbox.SelectedIndexChanged += new System.EventHandler(this.yeardropbox_SelectedIndexChanged);
            // 
            // importHTML
            // 
            this.importHTML.Location = new System.Drawing.Point(220, 351);
            this.importHTML.Margin = new System.Windows.Forms.Padding(2);
            this.importHTML.Name = "importHTML";
            this.importHTML.Size = new System.Drawing.Size(75, 24);
            this.importHTML.TabIndex = 23;
            this.importHTML.Text = "Import";
            this.importHTML.UseVisualStyleBackColor = true;
            this.importHTML.Click += new System.EventHandler(this.importHTML_Click);
            // 
            // AddGrade
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(362, 387);
            this.Controls.Add(this.importHTML);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.yeardropbox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.StudentNumber);
            this.Controls.Add(this.CourseTable);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "AddGrade";
            this.Text = "AddGrade";
            this.CourseTable.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TableLayoutPanel CourseTable;
        private System.Windows.Forms.RichTextBox richTextBox28;
        private System.Windows.Forms.RichTextBox richTextBox27;
        private System.Windows.Forms.RichTextBox richTextBox26;
        private System.Windows.Forms.RichTextBox richTextBox24;
        private System.Windows.Forms.RichTextBox richTextBox23;
        private System.Windows.Forms.RichTextBox richTextBox12;
        private System.Windows.Forms.RichTextBox richTextBox22;
        private System.Windows.Forms.RichTextBox richTextBox20;
        private System.Windows.Forms.RichTextBox richTextBox19;
        private System.Windows.Forms.RichTextBox richTextBox18;
        private System.Windows.Forms.RichTextBox richTextBox16;
        private System.Windows.Forms.RichTextBox richTextBox15;
        private System.Windows.Forms.RichTextBox richTextBox14;
        private System.Windows.Forms.RichTextBox richTextBox11;
        private System.Windows.Forms.RichTextBox richTextBox10;
        private System.Windows.Forms.RichTextBox richTextBox8;
        private System.Windows.Forms.RichTextBox richTextBox7;
        private System.Windows.Forms.RichTextBox richTextBox6;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.RichTextBox richTextBox4;
        private System.Windows.Forms.TextBox StudentNumber;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox yeardropbox;
        private System.Windows.Forms.Button importHTML;
    }
}