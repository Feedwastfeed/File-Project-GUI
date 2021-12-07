namespace File_Project_GUI
{
    partial class constraintsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(constraintsForm));
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.exitbtn = new System.Windows.Forms.Button();
            this.minibtn = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.notNullcheckbox = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.finishbtn = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.conditiontxtboxstring = new System.Windows.Forms.TextBox();
            this.conditiontxtboxidx = new System.Windows.Forms.TextBox();
            this.conditioncheckbox = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.uniquetxtboxidx = new System.Windows.Forms.TextBox();
            this.uniquecheckbox = new System.Windows.Forms.CheckBox();
            this.defaulttxtboxvalue = new System.Windows.Forms.TextBox();
            this.defaulttxtboxidx = new System.Windows.Forms.TextBox();
            this.defaultcheckbox = new System.Windows.Forms.CheckBox();
            this.notNulltxtboxidx = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.undobtn = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(1, 1);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(108, 26);
            this.label2.TabIndex = 11;
            this.label2.Text = "Constraints";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.exitbtn);
            this.panel1.Controls.Add(this.minibtn);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(657, 34);
            this.panel1.TabIndex = 12;
            // 
            // exitbtn
            // 
            this.exitbtn.BackgroundImage = global::File_Project_GUI.Properties.Resources.orange_power_button_md;
            this.exitbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.exitbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.exitbtn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.exitbtn.Location = new System.Drawing.Point(625, 2);
            this.exitbtn.Name = "exitbtn";
            this.exitbtn.Size = new System.Drawing.Size(25, 26);
            this.exitbtn.TabIndex = 14;
            this.exitbtn.UseVisualStyleBackColor = true;
            this.exitbtn.Click += new System.EventHandler(this.exitbtn_Click);
            // 
            // minibtn
            // 
            this.minibtn.BackgroundImage = global::File_Project_GUI.Properties.Resources._7b9fa5fc9d;
            this.minibtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.minibtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.minibtn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.minibtn.Location = new System.Drawing.Point(592, 3);
            this.minibtn.Name = "minibtn";
            this.minibtn.Size = new System.Drawing.Size(25, 25);
            this.minibtn.TabIndex = 15;
            this.minibtn.UseVisualStyleBackColor = true;
            this.minibtn.Click += new System.EventHandler(this.minibtn_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::File_Project_GUI.Properties.Resources.tools_and_utensils_block_lock_icon_3486_512x512;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(107, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(30, 28);
            this.pictureBox1.TabIndex = 13;
            this.pictureBox1.TabStop = false;
            // 
            // notNullcheckbox
            // 
            this.notNullcheckbox.AutoSize = true;
            this.notNullcheckbox.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold);
            this.notNullcheckbox.ForeColor = System.Drawing.Color.White;
            this.notNullcheckbox.Location = new System.Drawing.Point(37, 50);
            this.notNullcheckbox.Name = "notNullcheckbox";
            this.notNullcheckbox.Size = new System.Drawing.Size(150, 30);
            this.notNullcheckbox.TabIndex = 14;
            this.notNullcheckbox.Text = "Not NULL col.";
            this.notNullcheckbox.UseVisualStyleBackColor = true;
            this.notNullcheckbox.CheckedChanged += new System.EventHandler(this.notNullcheckbox_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.finishbtn);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.conditiontxtboxstring);
            this.groupBox1.Controls.Add(this.conditiontxtboxidx);
            this.groupBox1.Controls.Add(this.conditioncheckbox);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.uniquetxtboxidx);
            this.groupBox1.Controls.Add(this.uniquecheckbox);
            this.groupBox1.Controls.Add(this.defaulttxtboxvalue);
            this.groupBox1.Controls.Add(this.defaulttxtboxidx);
            this.groupBox1.Controls.Add(this.defaultcheckbox);
            this.groupBox1.Controls.Add(this.notNulltxtboxidx);
            this.groupBox1.Controls.Add(this.notNullcheckbox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold);
            this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.groupBox1.Location = new System.Drawing.Point(12, 51);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(633, 444);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Check your Cons.";
            // 
            // finishbtn
            // 
            this.finishbtn.BackgroundImage = global::File_Project_GUI.Properties.Resources.license_check;
            this.finishbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.finishbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.finishbtn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.finishbtn.Location = new System.Drawing.Point(301, 389);
            this.finishbtn.Name = "finishbtn";
            this.finishbtn.Size = new System.Drawing.Size(40, 40);
            this.finishbtn.TabIndex = 29;
            this.finishbtn.UseVisualStyleBackColor = true;
            this.finishbtn.Click += new System.EventHandler(this.finishbtn_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(32, 342);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(95, 26);
            this.label4.TabIndex = 28;
            this.label4.Text = "Condition";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(32, 288);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 26);
            this.label3.TabIndex = 27;
            this.label3.Text = "Columns";
            // 
            // conditiontxtboxstring
            // 
            this.conditiontxtboxstring.Enabled = false;
            this.conditiontxtboxstring.Location = new System.Drawing.Point(205, 347);
            this.conditiontxtboxstring.Name = "conditiontxtboxstring";
            this.conditiontxtboxstring.Size = new System.Drawing.Size(400, 27);
            this.conditiontxtboxstring.TabIndex = 26;
            // 
            // conditiontxtboxidx
            // 
            this.conditiontxtboxidx.Enabled = false;
            this.conditiontxtboxidx.Location = new System.Drawing.Point(205, 285);
            this.conditiontxtboxidx.Name = "conditiontxtboxidx";
            this.conditiontxtboxidx.Size = new System.Drawing.Size(400, 27);
            this.conditiontxtboxidx.TabIndex = 25;
            // 
            // conditioncheckbox
            // 
            this.conditioncheckbox.AutoSize = true;
            this.conditioncheckbox.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold);
            this.conditioncheckbox.ForeColor = System.Drawing.Color.White;
            this.conditioncheckbox.Location = new System.Drawing.Point(37, 236);
            this.conditioncheckbox.Name = "conditioncheckbox";
            this.conditioncheckbox.Size = new System.Drawing.Size(474, 30);
            this.conditioncheckbox.TabIndex = 24;
            this.conditioncheckbox.Text = "Check Condition  (let x be the number in condition)";
            this.conditioncheckbox.UseVisualStyleBackColor = true;
            this.conditioncheckbox.CheckedChanged += new System.EventHandler(this.conditioncheckbox_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(402, 113);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(60, 26);
            this.label7.TabIndex = 23;
            this.label7.Text = "Value";
            // 
            // uniquetxtboxidx
            // 
            this.uniquetxtboxidx.Enabled = false;
            this.uniquetxtboxidx.Location = new System.Drawing.Point(205, 180);
            this.uniquetxtboxidx.Name = "uniquetxtboxidx";
            this.uniquetxtboxidx.Size = new System.Drawing.Size(400, 27);
            this.uniquetxtboxidx.TabIndex = 22;
            // 
            // uniquecheckbox
            // 
            this.uniquecheckbox.AutoSize = true;
            this.uniquecheckbox.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold);
            this.uniquecheckbox.ForeColor = System.Drawing.Color.White;
            this.uniquecheckbox.Location = new System.Drawing.Point(38, 177);
            this.uniquecheckbox.Name = "uniquecheckbox";
            this.uniquecheckbox.Size = new System.Drawing.Size(130, 30);
            this.uniquecheckbox.TabIndex = 21;
            this.uniquecheckbox.Text = "Unique col.";
            this.uniquecheckbox.UseVisualStyleBackColor = true;
            this.uniquecheckbox.CheckedChanged += new System.EventHandler(this.uniquecheckbox_CheckedChanged);
            // 
            // defaulttxtboxvalue
            // 
            this.defaulttxtboxvalue.Enabled = false;
            this.defaulttxtboxvalue.Location = new System.Drawing.Point(483, 113);
            this.defaulttxtboxvalue.Name = "defaulttxtboxvalue";
            this.defaulttxtboxvalue.Size = new System.Drawing.Size(122, 27);
            this.defaulttxtboxvalue.TabIndex = 20;
            // 
            // defaulttxtboxidx
            // 
            this.defaulttxtboxidx.Enabled = false;
            this.defaulttxtboxidx.Location = new System.Drawing.Point(205, 113);
            this.defaulttxtboxidx.Name = "defaulttxtboxidx";
            this.defaulttxtboxidx.Size = new System.Drawing.Size(178, 27);
            this.defaulttxtboxidx.TabIndex = 19;
            // 
            // defaultcheckbox
            // 
            this.defaultcheckbox.AutoSize = true;
            this.defaultcheckbox.Font = new System.Drawing.Font("Calibri", 15.75F, System.Drawing.FontStyle.Bold);
            this.defaultcheckbox.ForeColor = System.Drawing.Color.White;
            this.defaultcheckbox.Location = new System.Drawing.Point(37, 113);
            this.defaultcheckbox.Name = "defaultcheckbox";
            this.defaultcheckbox.Size = new System.Drawing.Size(131, 30);
            this.defaultcheckbox.TabIndex = 18;
            this.defaultcheckbox.Text = "Default col.";
            this.defaultcheckbox.UseVisualStyleBackColor = true;
            this.defaultcheckbox.CheckedChanged += new System.EventHandler(this.defaultcheckbox_CheckedChanged);
            // 
            // notNulltxtboxidx
            // 
            this.notNulltxtboxidx.Enabled = false;
            this.notNulltxtboxidx.Location = new System.Drawing.Point(205, 50);
            this.notNulltxtboxidx.Name = "notNulltxtboxidx";
            this.notNulltxtboxidx.Size = new System.Drawing.Size(400, 27);
            this.notNulltxtboxidx.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Berlin Sans FB Demi", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(388, 423);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(244, 16);
            this.label1.TabIndex = 16;
            this.label1.Text = "Columns must be seperated by comma (,)";
            // 
            // undobtn
            // 
            this.undobtn.BackgroundImage = global::File_Project_GUI.Properties.Resources.undo_4_512;
            this.undobtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.undobtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.undobtn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.undobtn.Location = new System.Drawing.Point(5, 501);
            this.undobtn.Name = "undobtn";
            this.undobtn.Size = new System.Drawing.Size(30, 30);
            this.undobtn.TabIndex = 13;
            this.undobtn.UseVisualStyleBackColor = true;
            this.undobtn.Click += new System.EventHandler(this.undobtn_Click);
            // 
            // constraintsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(657, 534);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.undobtn);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "constraintsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "File Converter";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button exitbtn;
        private System.Windows.Forms.Button minibtn;
        private System.Windows.Forms.Button undobtn;
        private System.Windows.Forms.CheckBox notNullcheckbox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox notNulltxtboxidx;
        private System.Windows.Forms.TextBox defaulttxtboxvalue;
        private System.Windows.Forms.TextBox defaulttxtboxidx;
        private System.Windows.Forms.CheckBox defaultcheckbox;
        private System.Windows.Forms.TextBox uniquetxtboxidx;
        private System.Windows.Forms.CheckBox uniquecheckbox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox conditioncheckbox;
        private System.Windows.Forms.TextBox conditiontxtboxstring;
        private System.Windows.Forms.TextBox conditiontxtboxidx;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button finishbtn;
    }
}