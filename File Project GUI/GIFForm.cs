using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace File_Project_GUI
{
    public partial class GIFForm : Form
    {
        int choice;
        public GIFForm(Image img)
        {
            InitializeComponent();
            picbox.Image = img;
            picbox.SizeMode = PictureBoxSizeMode.StretchImage;
        }

    }
}
