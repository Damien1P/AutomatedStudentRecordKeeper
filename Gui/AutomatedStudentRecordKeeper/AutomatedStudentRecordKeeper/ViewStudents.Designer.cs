namespace AutomatedStudentRecordKeeper
{
    partial class ViewStudents
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
            this.removebutton = new System.Windows.Forms.Button();
            this.removetext = new System.Windows.Forms.TextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.studenttable = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.studenttable)).BeginInit();
            this.SuspendLayout();
            // 
            // removebutton
            // 
            this.removebutton.Location = new System.Drawing.Point(385, 480);
            this.removebutton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.removebutton.Name = "removebutton";
            this.removebutton.Size = new System.Drawing.Size(100, 28);
            this.removebutton.TabIndex = 15;
            this.removebutton.Text = "Remove";
            this.removebutton.UseVisualStyleBackColor = true;
            this.removebutton.Click += new System.EventHandler(this.removebutton_Click);
            // 
            // removetext
            // 
            this.removetext.Location = new System.Drawing.Point(249, 482);
            this.removetext.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.removetext.MaxLength = 7;
            this.removetext.Name = "removetext";
            this.removetext.Size = new System.Drawing.Size(132, 22);
            this.removetext.TabIndex = 16;
            this.removetext.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.removetext_KeyPress);
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(129, 486);
            this.label31.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(111, 17);
            this.label31.TabIndex = 17;
            this.label31.Text = "Student Number";
            // 
            // studenttable
            // 
            this.studenttable.AllowUserToAddRows = false;
            this.studenttable.AllowUserToDeleteRows = false;
            this.studenttable.AllowUserToOrderColumns = true;
            this.studenttable.AllowUserToResizeColumns = false;
            this.studenttable.AllowUserToResizeRows = false;
            this.studenttable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.studenttable.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.studenttable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.studenttable.Location = new System.Drawing.Point(52, 15);
            this.studenttable.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.studenttable.Name = "studenttable";
            this.studenttable.ReadOnly = true;
            this.studenttable.Size = new System.Drawing.Size(516, 458);
            this.studenttable.TabIndex = 34;
            this.studenttable.VirtualMode = true;
            // 
            // ViewStudents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 526);
            this.Controls.Add(this.studenttable);
            this.Controls.Add(this.label31);
            this.Controls.Add(this.removetext);
            this.Controls.Add(this.removebutton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "ViewStudents";
            this.Text = "ViewStudents";
            ((System.ComponentModel.ISupportInitialize)(this.studenttable)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button removebutton;
        private System.Windows.Forms.TextBox removetext;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.DataGridView studenttable;
    }
}