using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace File_Project_GUI
{
    public partial class constraintsForm : Form
    {
        Timer timer1 = new Timer();
        string value = "", con = "" , gifPath = Application.StartupPath.ToString().Remove(Application.StartupPath.ToString().Length-9,9) + "Properties\\Resources\\";
        bool [] error = new bool[4];
        bool [] funs = new bool[3];
        List<int> notNullidx = new List<int>();
        List<int> defaultidx = new List<int>();
        List<int> uniqueidx = new List<int>();
        List<int> conditionidx = new List<int>();


        public constraintsForm()
        {
            InitializeComponent();
            funs[0] = true;
            funs[1] = true;
            funs[2] = true;
        }


        //Exit btn 
        private void exitbtn_Click(object sender, EventArgs e)
        {
            Image image = Image.FromFile(gifPath + "end.png");
            GIFForm gify = new GIFForm(image);
            gify.Show();
            this.Hide();

            timer1.Interval = 4000;
            timer1.Tick += delegate
            {
                gify.Visible = false;
                Application.Exit();
            };
            timer1.Start();
        }


        //mini btn
        private void minibtn_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }


        //back
        private void undobtn_Click(object sender, EventArgs e)
        {
            startForm sform = new startForm();
            sform.Show();
            this.Hide();
        }


        //
        private void notNullcheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (notNullcheckbox.Checked)
            {
                notNulltxtboxidx.Enabled = true;
            }
            else if (notNullcheckbox.Checked == false)
            {
                notNulltxtboxidx.Enabled = false;
            }
        }


        //
        private void defaultcheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (defaultcheckbox.Checked)
            {
                defaulttxtboxidx.Enabled = true;
                defaulttxtboxvalue.Enabled = true;
            }
            else if (defaultcheckbox.Checked == false)
            {
                defaulttxtboxidx.Enabled = false;
                defaulttxtboxvalue.Enabled = false;
            }
        }


        //
        private void uniquecheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (uniquecheckbox.Checked)
            {
                uniquetxtboxidx.Enabled = true;
            }
            else if (uniquecheckbox.Checked == false)
            {
                uniquetxtboxidx.Enabled = false;
            }
        }


        //
        private void conditioncheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (conditioncheckbox.Checked)
            {
                conditiontxtboxidx.Enabled = true;
                conditiontxtboxstring.Enabled = true;
            }
            else if (conditioncheckbox.Checked == false)
            {
                conditiontxtboxidx.Enabled = false;
                conditiontxtboxstring.Enabled = false;
            }
        }


        //
        private void finishbtn_Click(object sender, EventArgs e)
        {

            ConstraintsFunctions cf = new ConstraintsFunctions();


            if (defaultcheckbox.Checked)
            {
                if (defaulttxtboxidx.Text != "" && defaulttxtboxvalue.Text != "")
                {
                    defaultidx = defaulttxtboxidx.Text.Split(',').Select(s => int.Parse(s)).ToList();
                    value = defaulttxtboxvalue.Text;
                    startForm.record = cf.Default(startForm.record, defaultidx, value);
                }
                else
                {
                    error[1] = true;
                }
            }


            if (notNullcheckbox.Checked)
            {
                if (notNulltxtboxidx.Text != "")
                {
                    notNullidx = notNulltxtboxidx.Text.Split(',').Select(s => int.Parse(s)).ToList();
                    funs[0] = cf.notNull(startForm.record, notNullidx);
                }
                else
                {
                    error[0] = true;
                }
            }


            if (uniquecheckbox.Checked)
            {
                if (uniquetxtboxidx.Text != "")
                {
                    uniqueidx = uniquetxtboxidx.Text.Split(',').Select(s => int.Parse(s)).ToList();
                    funs[1] = cf.Unique(startForm.record, uniqueidx);
                }
                else
                {
                    error[2] = true;
                }
            }


            if (conditioncheckbox.Checked)
            {
                if (conditiontxtboxidx.Text != "" && conditiontxtboxstring.Text != "")
                {
                    conditionidx = conditiontxtboxidx.Text.Split(',').Select(s => int.Parse(s)).ToList();
                    con = conditiontxtboxstring.Text;
                    funs[2] = cf.Condition(startForm.record, conditionidx, con);
                }
                else
                {
                    error[3] = true;
                }
            }


            if (error[0] || error[1] || error[2] || error[3])
            {
                MessageBox.Show("Please Fill the Information", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
      
        
            else
            {
                if (!funs[0] || !funs[1] || !funs[2])
                {
                    Image image = Image.FromFile(gifPath + "Error.gif");
                    GIFForm gify = new GIFForm(image);
                    gify.Show();
                    this.Hide();

                    timer1.Interval = 6000;
                    timer1.Tick += delegate
                    {
                        gify.Visible = false;
                        startForm st = new startForm();
                        st.Show();
                        timer1.Stop();
                    };
                    timer1.Start();

                }

                else
                {
                    startForm.fill_xml(startForm.xmlFileName, startForm.record);
                    Image image = Image.FromFile(gifPath + "Done.gif");
                    GIFForm gify = new GIFForm(image);
                    gify.Show();
                    this.Hide();
                    timer1.Interval = 5000;
                    timer1.Tick += delegate
                    {
                        gify.Visible = false;
                        startForm st = new startForm();
                        st.Show();
                        timer1.Stop();
                    };
                    timer1.Start();
                }

            }

        }
    }

}