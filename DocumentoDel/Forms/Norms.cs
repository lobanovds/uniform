using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentoDel.Forms
{
    public partial class Norms : Form
    {
        public DataTable NormsTable = new DataTable();

        public Norms()
        {
            InitializeComponent();
            textBox1.Text = DateTime.Now.Year.ToString();
            foreach (String tsi in new Forms.Data().Positions)
            comboBox1.Items.Add((Object)tsi);
            createNormsTable(NormsTable, new string[] { "ID", "YearOfNorms", "Position", "IsForWork", "NameUniform", "HowLongUseMonth", "CountInYear" });
            formatDatagrid(
                dataGridView1, 
                NormsTable,
                new string[] { "ID", "YearOfNorms", "Position", "IsForWork", "NameUniform", "HowLongUseMonth", "CountInYear" },
                new int[] { 60, 60, 130, 130, 330, 100, 50 },
                new string[] { "Уникальный номер норм", "ГодНорм", "Должность", "Выездной?", "Название СО из норм", "срок носки в мес.", "Штук в год" },
                new string[] { "ID", "ГодНорм", "Должность", "Выездной?", "Спецодежда", "Срок носки", "шт." },
                new bool[] { false, false, false, false, true, true, true });
        }

        private void createNormsTable(DataTable DT, string[] nameColumns)
        {
            DT.Reset();
            for (int i = 0; i < nameColumns.Length; i++)
                DT.Columns.Add(nameColumns[i]);
        }

        public void formatDatagrid(
            DataGridView DGV, 
            DataTable DT, 
            string[] nameColumns, 
            int[] widthColumns, 
            string[] tooltipColumns, 
            string[] captionColumns,
            bool[] visibleColumns)
        {
            DGV.DataSource = DT;
            for (int i = 0; i < nameColumns.Length; i++)
            {
                dataGridView1.Columns[nameColumns[i]].Width = widthColumns[i];
                dataGridView1.Columns[nameColumns[i]].ToolTipText = tooltipColumns[i];
                dataGridView1.Columns[nameColumns[i]].HeaderText = captionColumns[i];
                dataGridView1.Columns[nameColumns[i]].Visible = visibleColumns[i];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        public void SaveData()
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            XmlTextWriter xmlw1 = new XmlTextWriter(@"c:\shablon\data\norms.xml", UnicodeEncoding.UTF8);
            xmlw1.WriteStartDocument();
            xmlw1.WriteStartElement("Element");

            foreach (DataGridViewRow DR in dataGridView1.Rows)
            { 
                if 
            }
            
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                xmlw1.WriteStartElement("ROW");
                xmlw1.WriteAttributeString("ROW_ID", i.ToString());

                foreach (string textc in ColumnNames)
                {
                    if (textc != "HeIsMan")
                        xmlw1.WriteAttributeString(textc, dt.Rows[i][textc].ToString());
                    else
                        xmlw1.WriteAttributeString(textc, dt.Rows[i][textc].ToString() == "Муж" ? "true" : "false");
                }
                xmlw1.WriteFullEndElement();
            }
            xmlw1.WriteFullEndElement();
            xmlw1.WriteEndDocument();
            xmlw1.Close();
        }
    }
}
