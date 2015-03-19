using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace DocumentoDel
{
    public partial class Stations : Form
    {
        public DataTable DATA_STATIONS = new DataTable("DATASTATIONS");


        public Stations()
        {
            InitializeComponent();

            Create_Data_Stations();
            DATA_STATIONS.ReadXmlSchema(@"c:\shablon\data\ps.xslt");
            DATA_STATIONS.ReadXml(@"c:\shablon\data\ps.xml");
            comboBox1.SelectedIndex = 1;
            initdata();
        }

        private void Create_Data_Stations()
        {
            DATA_STATIONS.Reset();
            DATA_STATIONS.Columns.Add("RES");
            DATA_STATIONS.Columns.Add("PS");
            DATA_STATIONS.Columns.Add("PS_KV");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!(comboBox2.Text ==""))
            {
                bool none = true;
                foreach (DataRow dr in DATA_STATIONS.Rows)
                    if ((string)dr["PS"] == comboBox2.Text)
                    {
                        dr["PS_KV"] = textBox1.Text;
                        DATA_STATIONS.AcceptChanges();
                        none = false; 
                    }
                if (none)
                {
                    object[] rowadd = new object[3];
                    rowadd[0] = comboBox1.Text;
                    rowadd[1] = comboBox2.Text;
                    rowadd[2] = textBox1.Text==""?"_":textBox1.Text;
                    DATA_STATIONS.Rows.Add(rowadd);
                    DATA_STATIONS.AcceptChanges();
                    comboBox2.Items.Add(comboBox2.Text);
                }
            }
            DATA_STATIONS.WriteXml(@"c:\shablon\data\ps.xml");
            DATA_STATIONS.WriteXmlSchema(@"c:\shablon\data\ps.xslt");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DataRow dr in DATA_STATIONS.Rows)
                if ((string)dr["PS"] == comboBox2.Text)
                { dr.Delete(); }
            DATA_STATIONS.AcceptChanges();
            DATA_STATIONS.WriteXml(@"c:\shablon\data\ps.xml");
            DATA_STATIONS.WriteXmlSchema(@"c:\shablon\data\ps.xslt");
        }
        
        
        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            initdata();
        }

        private void initdata()
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            textBox1.Text = "";
            foreach (DataRow dr in DATA_STATIONS.Rows)
            { if ((string)dr["RES"] == comboBox1.Text) comboBox2.Items.Add((string)dr["PS"]); }
        }

        private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            textBox1.Text = "";
            foreach (DataRow dr in DATA_STATIONS.Rows)
            { if ((string)dr["PS"] == comboBox2.Text) textBox1.Text = (string)dr["PS_KV"]; }
        }

    }
}
