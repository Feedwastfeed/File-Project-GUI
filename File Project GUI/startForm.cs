using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace File_Project_GUI
{
    public partial class startForm : Form
    {
        public Timer displayTime = new Timer();
        public static string xmlFileName = "";
        string fileName = " ", content = "" ,gifPath = Application.StartupPath.ToString().Remove(Application.StartupPath.ToString().Length - 9, 9) + "Properties\\Resources\\";
        string dlrrow = "", dlrcol = "" , sheetNumber = "" ;
        public static List<List<string>> record = new List<List<string>>();
        public startForm()
        {
            InitializeComponent();
        }

        //filling the xml
        public static void fill_xml(string name, List<List<string>> rec)
        {
            name += ".xml";
            XmlWriter xml = XmlWriter.Create(name);

            xml.WriteStartDocument();
            xml.WriteStartElement("Table");

            for (int i = 0; i < rec.Count; i++)
            {
                xml.WriteStartElement("Row_" + Convert.ToString(i + 1));
                for (int j = 0; j < rec[i].Count; j++)
                {
                    xml.WriteStartElement("Col_" + Convert.ToString(j + 1));
                    xml.WriteString(rec[i][j]==null?"":rec[i][j]);
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
            xml.WriteEndDocument();

            xml.Close();
        }


        //exit button
        private void exitbtn_Click(object sender, EventArgs e)
        {
            Image image = Image.FromFile(gifPath + "end.png");
            GIFForm gify = new GIFForm(image);
            gify.Show();
            this.Hide();

            displayTime.Interval = 4000;
            displayTime.Tick += delegate
            {
                gify.Visible = false;
                Application.Exit();
            };
            displayTime.Start();

        }


        //mini button
        private void minibtn_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }


        //choose file type
        private void fromexcelrdobtn_CheckedChanged(object sender, EventArgs e)
        {
            if (fromexcelrdobtn.Checked)
            {
                fromexcelgroupbox.Visible = true;
                fromtxtgroupbox.Visible = false;
            }
        }


        private void fromtxtrdobtn_CheckedChanged(object sender, EventArgs e)
        {
            if (fromtxtrdobtn.Checked)
            {
                fromexcelgroupbox.Visible = false;
                fromtxtgroupbox.Visible = true;
            }
        }


        //pick txt file
        private void browseFilebtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Text files |*.txt";
            if (fileDialog.ShowDialog() == DialogResult.OK )
            {
                if (fileDialog.FileName != null && fileDialog.FileName != "")
                {
                    fileName = fileDialog.FileName;
                    dlrrowtxtbox.Enabled = true;
                    dlrcoltxtbox.Enabled = true;
                    importFilebtn.Enabled = true;
                    xmlFileNametxtboxT.Enabled = true;
                }
            }
            else
            {
                Image image = Image.FromFile(gifPath + "wrong.gif");
                GIFForm gify = new GIFForm(image);
                gify.Show();

                displayTime.Interval = 3000;
                displayTime.Tick += delegate
                {
                    gify.Visible = false;
                };
                displayTime.Start();

            }
                
            
        }


        //pick excel file
        private void browseFileexcelbtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Text files |*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                if (fileDialog.FileName != null && fileDialog.FileName != "")
                {
                    fileName = fileDialog.FileName;
                    sheetNumbertxtbox.Enabled = true;
                    importexcelbtn.Enabled = true;
                    xmlFileNametxtboxE.Enabled = true;
                }
            }
            else
            {
                Image image = Image.FromFile(gifPath + "wrong.gif");
                GIFForm gify = new GIFForm(image);
                gify.Show();

                displayTime.Interval = 3000;
                displayTime.Tick += delegate
                {
                    gify.Visible = false;
                };
                displayTime.Start();

            }
        }


        //convertion 
        private void importexcelbtn_Click(object sender, EventArgs e)
        {
            if (sheetNumbertxtbox.Text != "" && xmlFileNametxtboxE.Text !="")
            {
                sheetNumber = sheetNumbertxtbox.Text;
                xmlFileName = xmlFileNametxtboxE.Text;
              
                //import from excel
                Excel excel = new Excel(fileName,int.Parse(sheetNumber));
                record = excel.readAll();
                excel.Close();
                Image image = Image.FromFile(gifPath + "choose.gif");
                GIFForm gify = new GIFForm(image);
                gify.Show();
                this.Hide();

                displayTime.Interval = 13000;
                displayTime.Tick += delegate
                {
                    gify.Visible = false;
                    constraintsForm cform = new constraintsForm();
                    cform.Show();
                    displayTime.Stop();
                };
                displayTime.Start();
                
                
            }
            else
            {
                MessageBox.Show("Please Enter the Sheet Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void importFilebtn_Click(object sender, EventArgs e)
        {
            if (dlrcoltxtbox.Text != "" && dlrrowtxtbox.Text != "") 
            {
                dlrrow = dlrrowtxtbox.Text;
                dlrcol = dlrcoltxtbox.Text;
                xmlFileName = xmlFileNametxtboxT.Text;
                
                //import from text
                content = File.ReadAllText(fileName);

                foreach (var row in content.Split(Convert.ToChar(dlrrow)))
                {
                    record.Add(row.Split(Char.Parse(dlrcol)).ToList<string>());
                }

                Image image = Image.FromFile(gifPath + "choose.gif");
                GIFForm gify = new GIFForm(image);
                gify.Show();
                this.Hide();

                displayTime.Interval = 13000;
                displayTime.Tick += delegate
                {
                    gify.Visible = false;
                    constraintsForm cform = new constraintsForm();
                    cform.Show();
                    displayTime.Stop();
                };
                displayTime.Start();
                

            }
            else
            {
                MessageBox.Show("Please Enter the delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
