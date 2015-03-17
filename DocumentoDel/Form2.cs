using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentoDel.BLL;
using DocumentoDel.IDAL;
using DocumentoDel.Properties;
using DocumentoDel.MSWord;

namespace DocumentoDel
{
    public partial class Form2 : Form
    {
        public Word word;
        public Form2()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            word = new Word();
        }
    }
}
