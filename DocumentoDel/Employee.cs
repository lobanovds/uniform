using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentoDel
{
    public partial class Employee : Form
    {

        public struct SizeOfEmployeeByYear
        {
            public int Tabel;
            public DateTime MeasureDate;
            public int Length; //Высота из классификатора. По ним будем брать ростовки
            public int Wide; //Ширина из классификатора. По ним будем брать размеры одежды
            public int BootsSummer; //Размер ботинок лето
            public int BootsWinter; //Размер ботинок зима
            public int Head; //Размер головы
        }

        public struct Employees
        {
            public int Tabel;
            public bool HeIsMan;
            public String LastName;
            public String Name;
            public String SurName;
            public DateTime Birthday;
            public DateTime StartWorking;
            public int PositionByTabel;
        }

        public struct PositionChange
        {
            public int Tabel;
            public DateTime Start;
            public DateTime End;
            public int positionID;
            public Boolean isNow;
        }

        public struct Position
        {
            public int positionID;
            public String positionName;
            public String SubPositionName;
        }

        
        public DataTable PositionTable = new DataTable("Position");
        public DataTable PositionChangeTable = new DataTable("PositionChange");
        public DataTable EmployeesTable = new DataTable("Employees");
        public DataTable SizeOfEmployeeByYearTable = new DataTable("SizeOfEmployeeByYear");

        public Employee()
        {
            InitializeComponent();
            comboBox1.Text = DateTime.Now.Year.ToString();
            createEmployeesTable();
            SvidDataGridInit();
        }

        //TODO: Добавить в таблицу с должностями
        /// <summary>
        /// Создаём таблицу для Сотрудников
        /// </summary>
        private void createEmployeesTable()
        {
            EmployeesTable.Reset();
            EmployeesTable.Columns.Add("Tabel");
            EmployeesTable.Columns.Add("HeIsMan");
            EmployeesTable.Columns.Add("LastName");
            EmployeesTable.Columns.Add("Name");
            EmployeesTable.Columns.Add("SurName");
            EmployeesTable.Columns.Add("Birthday");
            EmployeesTable.Columns.Add("StartWorking");
            EmployeesTable.Columns.Add("PositionByTabel");
            EmployeesTable.AcceptChanges();
        }

        /// <summary>
        /// Создаём таблицу для Должностей
        /// </summary>
        private void createPositionTable()
        {
            PositionTable.Reset();
            PositionTable.Columns.Add("positionID");
            PositionTable.Columns.Add("positionName");
            PositionTable.Columns.Add("SubPositionName");
            PositionTable.AcceptChanges();
        }


        /// <summary>
        /// Оформление грида (таблицы для сотрудников)
        /// </summary>
        private void SvidDataGridInit()
        {
            dataGridView1.DataSource = EmployeesTable;
            dataGridView1.Columns["Tabel"].Width = 60;
            dataGridView1.Columns["HeIsMan"].Width = 60;
            dataGridView1.Columns["LastName"].Width = 130;
            dataGridView1.Columns["Name"].Width = 130;
            dataGridView1.Columns["Surname"].Width = 130;
            dataGridView1.Columns["Birthday"].Width = 100;
            dataGridView1.Columns["StartWorking"].Width = 80;
            dataGridView1.Columns["PositionByTabel"].Width = 80;

            dataGridView1.Columns["Tabel"].ToolTipText = "Табель";
            dataGridView1.Columns["HeIsMan"].ToolTipText = "Пол";
            dataGridView1.Columns["LastName"].ToolTipText = "Фамилия";
            dataGridView1.Columns["Name"].ToolTipText = "Имя";
            dataGridView1.Columns["Surname"].ToolTipText = "Отчество";
            dataGridView1.Columns["Birthday"].ToolTipText = "День рождения";
            dataGridView1.Columns["StartWorking"].ToolTipText = "Дата начала работы";
            dataGridView1.Columns["PositionByTabel"].ToolTipText = "Должность";

            dataGridView1.Columns["Tabel"].HeaderText = "Таб.";
            dataGridView1.Columns["HeIsMan"].HeaderText = "Пол";
            dataGridView1.Columns["LastName"].HeaderText = "Фамилия";
            dataGridView1.Columns["Name"].HeaderText = "Имя";
            dataGridView1.Columns["Surname"].HeaderText = "Отчество";
            dataGridView1.Columns["Birthday"].HeaderText = "ДР";
            dataGridView1.Columns["StartWorking"].HeaderText = "Устроился";
            dataGridView1.Columns["PositionByTabel"].HeaderText = "Должность";

            /*
            dataGridView1.Columns["Tabel"]
            dataGridView1.Columns["HeIsMan"]
            dataGridView1.Columns["LastName"]
            dataGridView1.Columns["Name"]
            dataGridView1.Columns["Surname"]
            dataGridView1.Columns["Birthday"]
            dataGridView1.Columns["StartWorking"]
             * 
             * .ToolTipText = "";
             * .HeaderText = "";
             */
        }

    }
}
