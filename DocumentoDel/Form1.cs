using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using DocumentoDel.BLL;
using System.Reflection;
using System.Threading;
using System.Xml;
using System.Text;
using System.Net;
using System.Globalization;


namespace DocumentoDel
{
    public partial class Form1 : Form
    {
        public string pathmano = @"c:\shablon\temp\mano\";
        public object wordapp;
        public object worddocument;

        public CreateAkt CreateAktClass;
        public Debug DebugClass = new Debug();

        //Рабочие таблицы
        public DataTable DT_SHETCHIK = new DataTable("IDCounterTable");
        public DataTable DT_SVID_TT = new DataTable("IDSVIDTT");
        public DataTable DT_MANOMETR = new DataTable("IDMANOMETR");
        public DataTable DT_PODSTATIONS = new DataTable("DATASTATIONS");

        //Здесь хранятся ФИО проверяющих
        public String[] Control1 = new String[2];
        public String[] Control2 = new String[3];

        //Здесь хранятся данные о калибровщиках
        public String[] Calibrator = new String[2];

        public int column = 0;
        public int row = 0;
        public int lastid = 0;

        Int32 press; //давление
        Int32 temperature; //температура
        Int32 humidity; //влажность
        Int32 getpress; //давление
        Int32 gettemperature; //температура
        Int32 gethumidity; //влажность

        //Включить модуль свидетельств для трансформаторов тока?
        public Boolean tt_ON = false;

        //Включить модуль свидетельств для ТС и Манометров?
        public Boolean mano_ON = true;

        private void checkpath()
        {
            if (!Directory.Exists(pathmano)) Directory.CreateDirectory(pathmano);
        }

        public Form1()
        {
            InitializeComponent();
            checkpath();
            textBox_manoSvidDate.Text = DateTime.Now.ToShortDateString();
            //В случае изменения числа проверяющих с той 
            //или иной стороны - необходима переинициализация Control[!] строки 29-30
            //Вводим данные о проверяющих из ЦСМ
            Control1[0] = "Сухинова Л.Г.";
            Control1[1] = "Демченко Е.В.";
            //Вводим данные о проверяющих из нашей организации
            Control2[0] = "Канке А.А.";
            Control2[1] = "Балай А.И.";
            Control2[2] = "Долгих Е.В.";
            //Калибровщики
            Calibrator[0] = "Корниенков Е.Д.";
            Calibrator[1] = "Лобанов Д.С.";
            

            //Создаём класс для обращения к Word
            CreateAktClass = new CreateAkt();
            
            //Проверка ошибок
            if (!(System.IO.Directory.Exists(@"C:\shablon\")))
            { 
                MessageBox.Show(@"Внимание! Директория C:\shablon\ отсутствует");
            }
            else if (!(System.IO.Directory.Exists(@"C:\shablon\data\")))
            {
                MessageBox.Show(@"Внимание! Директория C:\shablon\data\ отсутствует");
            }
            CreateAktClass.path_filename = @"C:\shablon\data\";

            //Создание структуры таблицы для счётчиков
            createDT_SCHETCHIK();
            
            //Оформляем таблицу в понятную структуру
            SchetchikDataGridInit();
            
            //если у нас есть ещё и таблицы для ТТ, то оформляем таблицу во второй вкладке
            if (tt_ON)
            {
                //Создание структуры таблицы для трансформаторов
                createDT_SVID_TT();
                
                //Настраиваем внешний вид таблицы
                SvidDataGridInit();
            }

            if (mano_ON)
            {
                //Создание структуры таблицы для ТС и Манометров
                createDT_MANOMETR();

                //Настраиваем внешний вид таблицы
                ManoDataGridInit();

                comboBox1.SelectedIndex = 0;
            }

            //Вывод даты во втором текстбоксе
            textBox_date1.Text = DateTime.Now.Date.ToShortDateString();

            //если дата вида 25.06.13, то вставляем 20 в формат года
            textBox_date1.Text = textBox_date1.Text.Length == 8 ? textBox_date1.Text.Insert(6, "20") : textBox_date1.Text; 
            //если вместо точек запятые, меняем запятые на точки
            textBox_date2.Text = textBox_date1.Text = textBox_date1.Text.Replace(",",".");

            //Сразу определяем путь для сохранения
            StreamReader strReader = new StreamReader(@"c:\shablon\data\path.txt");

            textBox_path.Text = strReader.ReadLine();
            strReader.Close();

            //устанавливаем папку по умолчанию как ту, что указана в выборе директории
            folderBrowserDialog1.SelectedPath = @textBox_path.Text;

            if (tt_ON)
            {
                getMeteo();
                insertXML();
            }
            else { tabControl1.TabPages["TabPageTT"].Dispose(); }

            if (mano_ON){}
            else { tabControl1.TabPages["TabPageMano"].Dispose(); }

            Create_Data_Stations();
            DT_PODSTATIONS.ReadXmlSchema(@"c:\shablon\data\ps.xslt");
            DT_PODSTATIONS.ReadXml(@"c:\shablon\data\ps.xml");
            DT_PODSTATIONS.AcceptChanges();
        }


        /// <summary>
        /// Создаём таблицу для подстанций
        /// </summary>
        private void Create_Data_Stations()
        {
            DT_PODSTATIONS.Reset();
            DT_PODSTATIONS.Columns.Add("RES");
            DT_PODSTATIONS.Columns.Add("PS");
            DT_PODSTATIONS.Columns.Add("PS_KV");
        }


        /// <summary>
        /// Создаём таблицу для счётчиков
        /// </summary>
        private void createDT_SCHETCHIK()
        {
            DT_SHETCHIK.Reset();
            DT_SHETCHIK.Columns.Add("ROW_ID");
            DT_SHETCHIK.Columns.Add("ID");
            DT_SHETCHIK.Columns.Add("Poverka");
            DT_SHETCHIK.Columns.Add("Tariff1");
            DT_SHETCHIK.Columns.Add("Tariff2");
            DT_SHETCHIK.Columns.Add("Kvartal");
            DT_SHETCHIK.AcceptChanges();
        }

        /// <summary>
        /// Оформление грида (таблицы для счётчиков)
        /// </summary>
        private void SchetchikDataGridInit()
        {
            dataGridView1.DataSource = DT_SHETCHIK;
            dataGridView1.Columns["ROW_ID"].Visible = false;
            dataGridView1.Columns["ID"].Width = 80;
            dataGridView1.Columns["Poverka"].Width = 80;
            dataGridView1.Columns["Tariff1"].Width = 80;
            dataGridView1.Columns["Tariff2"].Width = 80;
            dataGridView1.Columns["Kvartal"].Width = 80;
            dataGridView1.Columns["ID"].HeaderText = "№Счётчика";
            dataGridView1.Columns["Poverka"].HeaderText = "Поверка";
            dataGridView1.Columns["Poverka"].ToolTipText = "Дата поверки";
            dataGridView1.Columns["Tariff1"].HeaderText = "Тариф1";
            dataGridView1.Columns["Tariff2"].HeaderText = "Тариф2";
            dataGridView1.Columns["Tariff2"].ToolTipText = "Для двухтарифных счётчиков";
            dataGridView1.Columns["Kvartal"].HeaderText = "Квартал";
        }

        /// <summary>
        /// Создаём таблицу для свидетельств ТТ
        /// </summary>
        private void createDT_SVID_TT()
        {
            DT_SVID_TT.Reset();
            DT_SVID_TT.Columns.Add("TYPE");
            DT_SVID_TT.Columns.Add("ID");
            DT_SVID_TT.Columns.Add("Power");
            DT_SVID_TT.Columns.Add("KTT");
            DT_SVID_TT.Columns.Add("Interval");
            DT_SVID_TT.AcceptChanges();
        }

        /// <summary>
        /// Оформление грида (таблицы для ТТ)
        /// </summary>
        private void SvidDataGridInit()
        {
            dataGridView2.DataSource = DT_SVID_TT;
            dataGridView2.Columns["TYPE"].Width = 100;
            dataGridView2.Columns["ID"].Width = 80;
            dataGridView2.Columns["Power"].Width = 80;
            dataGridView2.Columns["KTT"].Width = 80;
            dataGridView2.Columns["TYPE"].HeaderText = "Тип";
            dataGridView2.Columns["TYPE"].ToolTipText = "Тип трансформатора тока";
            dataGridView2.Columns["ID"].HeaderText = "Номер";
            dataGridView2.Columns["ID"].ToolTipText = "Номер трансформатора тока";
            dataGridView2.Columns["Power"].HeaderText = "Нагрузка ВА";
            dataGridView2.Columns["Power"].ToolTipText = "Нагрузка ВА";
            dataGridView2.Columns["KTT"].HeaderText = "Номинальный ток";
            dataGridView2.Columns["KTT"].ToolTipText = "Номинальный первичный ток";
            dataGridView2.Columns["Interval"].HeaderText = "Межпов. Интервал";
            dataGridView2.Columns["Interval"].ToolTipText = "Межповерочный интервал";
        }

        /// <summary>
        /// Создаём таблицу для приборов(манометров и ТКП)
        /// </summary>
        private void createDT_MANOMETR()
        {
            DT_MANOMETR.Reset();
            DT_MANOMETR.Columns.Add("ID_SVID");
            DT_MANOMETR.Columns.Add("MANO_NAME");
            DT_MANOMETR.Columns.Add("MANO_ID");
            DT_MANOMETR.Columns.Add("MANO_DATE");
            DT_MANOMETR.Columns.Add("MANO_LENGTH");
            DT_MANOMETR.Columns.Add("SCALE");
            DT_MANOMETR.Columns.Add("PLACE_RES");
            DT_MANOMETR.Columns.Add("PLACE_STATION");
            DT_MANOMETR.Columns.Add("PLACE");
            DT_MANOMETR.AcceptChanges();
        }

        /// <summary>
        /// Оформление грида (таблицы для Манометров)
        /// </summary>
        private void ManoDataGridInit()
        {
            dataGridView3.DataSource = DT_MANOMETR;
            dataGridView3.Columns["ID_SVID"].Width = 60;
            dataGridView3.Columns["ID_SVID"].ReadOnly = true;
            dataGridView3.Columns["MANO_NAME"].Width = 70;
            dataGridView3.Columns["MANO_NAME"].ContextMenuStrip = contextMenuStrip3;
            dataGridView3.Columns["MANO_ID"].Width = 70;
            dataGridView3.Columns["MANO_ID"].ContextMenuStrip = contextMenuStripNumbersNotExist;
            dataGridView3.Columns["MANO_DATE"].Width = 70;
            dataGridView3.Columns["MANO_DATE"].ContextMenuStrip = ContextYears;
            dataGridView3.Columns["MANO_LENGTH"].Width = 50;
            dataGridView3.Columns["SCALE"].Width = 60;
            dataGridView3.Columns["PLACE_RES"].Width = 80;
            dataGridView3.Columns["PLACE_STATION"].Width = 80;
            dataGridView3.Columns["PLACE"].Width = 80;
            dataGridView3.Columns["ID_SVID"].HeaderText = "Св-во";
            dataGridView3.Columns["ID_SVID"].ToolTipText = "Номер свидетельства";
            dataGridView3.Columns["MANO_NAME"].HeaderText = "Тип прибора";
            dataGridView3.Columns["MANO_NAME"].ToolTipText = "Тип манометра или термосигнализатора";
            dataGridView3.Columns["MANO_ID"].HeaderText = "Номер прибора";
            dataGridView3.Columns["MANO_ID"].ToolTipText = "Номер манометра или термосигнализатора";
            dataGridView3.Columns["MANO_DATE"].HeaderText = "Год выпуска";
            dataGridView3.Columns["MANO_DATE"].ToolTipText = "Год выпуска";
            dataGridView3.Columns["MANO_LENGTH"].HeaderText = "Длина";
            dataGridView3.Columns["MANO_LENGTH"].ToolTipText = "Длина капиляра термосигнализатора";
            dataGridView3.Columns["MANO_LENGTH"].ContextMenuStrip = contextMenuStripLength;
            dataGridView3.Columns["SCALE"].HeaderText = "Диапазон";
            dataGridView3.Columns["SCALE"].ToolTipText = "Диапазон измерений";
            dataGridView3.Columns["SCALE"].ContextMenuStrip = contextMenuStripDiapazon;
            dataGridView3.Columns["PLACE_RES"].HeaderText = "РЭС";
            dataGridView3.Columns["PLACE_RES"].ToolTipText = "Район Электрических Сетей, где установлен прибор";
            dataGridView3.Columns["PLACE_RES"].ContextMenuStrip = contextMenuStrip1;
            dataGridView3.Columns["PLACE_STATION"].HeaderText = "ПС";
            dataGridView3.Columns["PLACE_STATION"].ToolTipText = "Подстанция, где установлен прибор";
            dataGridView3.Columns["PLACE_STATION"].ContextMenuStrip = contextMenuStrip2;
            dataGridView3.Columns["PLACE"].HeaderText = "Место";
            dataGridView3.Columns["PLACE"].ToolTipText = "Место, где установлен прибор";

        }



        /// <summary>
        /// Получаем значение давления с сайта гисметео
        /// </summary>
        /// <returns>возвращается давление в мм.рт.ст.</returns>
        public Int32 getMeteo() 
        {
            try
            {
                System.Net.WebClient WC = new System.Net.WebClient();
                if (!File.Exists(@"C:\shablon\Testing\meteodata.xml"))
                    File.Copy(@"C:\shablon\Testing\meteodata.xml", @"C:\shablon\Testing\meteodatatemp.xml");
                else
                    File.Replace(@"C:\shablon\Testing\meteodata.xml", @"C:\shablon\Testing\meteodatatemp.xml","temp1");
                WC.DownloadFile("http://informer.gismeteo.ru/rss/23471.xml", @"C:\shablon\Testing\meteodata.xml");
            }
            catch {
                if (!File.Exists(@"C:\shablon\Testing\meteodata.xml"))
                    File.Copy(@"C:\shablon\Testing\meteodatatemp.xml", @"C:\shablon\Testing\meteodata.xml");
                else
                    File.Replace(@"C:\shablon\Testing\meteodatatemp.xml", @"C:\shablon\Testing\meteodata.xml", "temp2");
            }
            XmlDocument xmdoc = new XmlDocument();
            xmdoc.Load(@"C:\shablon\Testing\meteodata.xml");
            String data = xmdoc.DocumentElement.SelectNodes("channel/item/description")[0].InnerText;
            data.Substring(data.IndexOf("давление ") + 9, 3);
            Int32 pressure = Convert.ToInt32(data.Substring(data.IndexOf("давление ") + "давление ".Count(), 3)) + 1;
            
            string gettemperature_string="0";
            try
            {
                gettemperature_string = data.Substring(data.IndexOf("температура ") + "температура ".Count(), 3);
                gettemperature_string = gettemperature_string.Remove(gettemperature_string.IndexOf("."));
                gettemperature_string = gettemperature_string.Remove(gettemperature_string.IndexOf("."));
            }
            catch { }
            int tempTemp = Convert.ToInt32(gettemperature_string);
            gettemperature = tempTemp == 0 ? 0 : tempTemp > 0 ? tempTemp + 1 : tempTemp - 1;
            
            return pressure;
        }

        /// <summary>
        /// Вставляем данные в наш XML-файл (дата, температура, влажность, давление)
        /// </summary>
        public void insertXML()
        {
            dataSet1.Tables.Add("row");
            dataSet1.Tables[0].Columns.Add("date");
            dataSet1.Tables[0].Columns.Add("temperature");
            dataSet1.Tables[0].Columns.Add("humidity");
            dataSet1.Tables[0].Columns.Add("pressure");
            try
            {
                dataSet1.ReadXml(@"C:\shablon\Testing\data.xml", XmlReadMode.Auto);
            }
            catch (Exception e)
            {
                string alert = e.ToString();
                XmlTextWriter xmlw1 = new XmlTextWriter(@"C:\shablon\Testing\data.xml", UnicodeEncoding.UTF8);
                xmlw1.WriteStartDocument(true);
                xmlw1.WriteStartElement("XML");
                xmlw1.WriteFullEndElement();
                xmlw1.WriteEndDocument();
                xmlw1.Close();
            }

            if (dataSet1.Tables[0].Rows.Count > 0 && dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1][0].ToString() == DateTime.Now.ToShortDateString())
            {
                press = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["pressure"]);
                label9.Text = "Давление: " + press.ToString() + " мм.рт.ст.";
                humidity = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["humidity"]);
                label10.Text = "Влажность: " + humidity.ToString();
                temperature = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["temperature"]);
                label11.Text = "Температура: " + temperature.ToString() + " C ";

            }
            else
            {
                XmlDocument xmdoc = new XmlDocument();
                xmdoc.Load(@"C:\shablon\Testing\meteodata.xml");
                String data = xmdoc.DocumentElement.SelectNodes("channel/item/description")[0].InnerText;
                data.Substring(data.IndexOf("давление ") + 9, 3);
                press = Convert.ToInt32(data.Substring(data.IndexOf("давление ") + 9, 3)) + 1;

                object[] newstring = new object[4];
                newstring[0] = (object)DateTime.Now.ToShortDateString();
                newstring[1] = (object)gettemperature;
                newstring[2] = (object)4;//принудительно вставили влажность =4
                newstring[3] = (object)getMeteo();
                dataSet1.Tables[0].Rows.Add(newstring);
                dataSet1.WriteXml(@"C:\shablon\Testing\data.xml");

                //Заполняем данные в форме
                press = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["pressure"]);
                label9.Text = "Давление:" + press.ToString();
                humidity = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["humidity"]);
                label10.Text = "Влажность:" + humidity.ToString();
                temperature = Convert.ToInt32(dataSet1.Tables[0].Rows[dataSet1.Tables[0].Rows.Count - 1]["temperature"]);
                label11.Text = "Температура:" + temperature.ToString() + " C";
            }
        }


        /// <summary>
        /// Оформляем правильно выпадающий список 
        /// с проверяющими протокол в зависимости от 
        /// значения CheckBox "Наш проверяющий"
        /// </summary>
        private void ComboBoxInit()
        {
            String[] Control3 = checkBox1.Checked ? Control2 : Control1;
            for (int i = 0; i < Control3.Length; i++)
            {
                comboBox_control.Items.Insert(i, (object)(Control3[i]));
            }
            for (int i = 0; i < Control1.Length; i++)
            {
                comboBox2.Items.Insert(i, (object)(Control1[i]));
            }

            comboBox_control.Text = checkBox1.Checked ? "Канке А.А." : "Сухинова Л.Г.";
            comboBox_control.SelectedIndex = 0;
        }


        

        /// <summary>
        /// Жмём на кнопку, создаём поверочные документы
        /// </summary>
        private void submitButton_Click(object sender, EventArgs e)
        {
            DataTable NDT = (DataTable)dataGridView1.DataSource;
            NDT.TableName = "POVERKA";
            SaveData(NDT);

            string NumberCounter = "";
            string Poverka = "";
            string Kvartal = "";
            string tempstring = "";
            
            if (comboBox_maker.Text == "Выполнил")
            {
                MessageBox.Show("Выберите выполнившего поверку");
                comboBox_maker.Select();
            }
            else if (comboBox_control.Text == "Проверил")
            {
                MessageBox.Show("Выберите проверяющего");
                comboBox_control.Select();
            }
            else
            {
                for (int i = 0; i < NDT.Rows.Count; i++)
                {
                    CreateAktClass.openSchet(comboBox1_schetchikType.Text); //открываем тот или иной шаблон

                    if ((NDT.Rows[i]["Poverka"].ToString() == "") && (textBox_datePoverka.Text == ""))
                        dataGridView1.Rows[i].Selected = true;
                                        
                    NumberCounter = NDT.Rows[i]["ID"].ToString();
                    changer("Nomer_schetchika", NumberCounter);     //Номер счётчика

                    changer("vipolnil", comboBox_maker.Text);       //Кто составил протокол

                    changer("proveril_state", checkBox1.Checked ?
                        "Проверил инженер ЭТЛ                                            " :
                        "Проверил ФБУ «Тюменский ЦСМ»                         "); 
                    changer("proveril", comboBox_control.Text);     //Кто проверил протокол
                    
                    String ToDate = ((textBox_date1.Text.Length != 10) || (textBox_date1.Text.Length != 8)) ? DateTime.Now.Date.ToShortDateString() : textBox_date1.Text;
                    ToDate = ToDate.Length == 8 ? ToDate.Insert(6, "20") : ToDate;
                    changer("date", ToDate);                //Дата заполнения протокола

                    Poverka = (NDT.Rows[i]["Poverka"].ToString() == "") ? textBox_datePoverka.Text == "" ? "0" : textBox_datePoverka.Text : NDT.Rows[i]["Poverka"].ToString();
                    Poverka = Poverka.Length == 8 ? Poverka.Insert(6, "20") : Poverka;
                    Poverka = Poverka == "0" ? DateTime.Now.ToShortDateString() : Poverka;
                    
                    changer("datavipuska", Poverka);        //Дата выпуска счётчика
                    changer("godpoverki", Poverka.Substring(6, 4));    //Год поверки

                    
                    if (NDT.Rows[i]["Kvartal"].ToString().Contains("I"))
                        Kvartal = NDT.Rows[i]["Kvartal"].ToString();
                    else
                        Kvartal = PoverkaKvartal(Poverka);
                    
                    changer("kvartal", Kvartal);     //Квартал поверки
                    
                    Randomizer();

                    string SummerTimeText1 = radioButton1.Checked ? "летнее" : "зимнее";
                    string SummerTimeText2 = radioButton1.Checked ? "зимнее" : "летнее";
                    string allowtime = checkBox_allowtime.Checked ? "разрешён" : "запрещён";
                    changer("letozima2", SummerTimeText1);       //Время счётчика
                    changer("letozima1", SummerTimeText2);       //Переход на другое время
                    changer("rule", allowtime);         //Разрешено ли переходить на другое время

                    string tarif1 = NDT.Rows[i]["Tariff1"].ToString();
                    string tarif2 = NDT.Rows[i]["Tariff2"].ToString();
                    //string tarif_summ; 

                    ///параметры счётчиков, которые отличаются
                    #region counters_differents
                    switch (comboBox1_schetchikType.Text)
                    {
                        case "1т. СЭБ2а (100A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length - 3, 3));     //Связной счётчика
                                tempstring = "-seb2a(1f-1t)-100A";
                                break;
                            }
                        case "2т. СЭБ2а (100A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length - 3, 3));     //Связной счётчика
                                tempstring = "-seb2a(1f-2t)-100A";
                                break;
                            }
                        case "1т. СЭБ2а (50A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length - 3, 3));     //Связной счётчика
                                tempstring = "-seb2a(1f-1t)-50A";
                                break;
                            }
                        case "2т. СЭБ2а (50A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length - 3, 3));     //Связной счётчика
                                tempstring = "-seb2a(1f-2t)-50A";
                                break;
                            }
                        case "1т. Маяк (80A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length-3, 3));     //Связной счётчика
                                tempstring = "-mayak(1f-1t)";
                                break;
                            }
                        case "2т. Маяк (80A)":
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length-3, 3));     //Связной счётчика
                                tempstring = "-mayak(1f-2t)";
                                break;
                            }
                        case "1т. Меркурий  (60A)":
                            {
                                ChangeTar6(ref tarif1);
                                ChangeTar6(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length-6, 6));     //Связной счётчика
                                tempstring = "-merkurii(1f-1t)";
                                break;
                            }
                        case "2т. Меркурий  (60A)":
                            {
                                ChangeTar6(ref tarif1);
                                ChangeTar6(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length-6, 6));     //Связной счётчика
                                tempstring = "-merkurii(1f-2t)";
                                break;
                            }
                        default:
                            {
                                ChangeTar5(ref tarif1);
                                ChangeTar5(ref tarif2);
                                changer("svyaznoy", NumberCounter.Substring(NumberCounter.Length - 3, 3));     //Связной счётчика
                                tempstring = "-seb2a(1f-1t)";
                                break;
                            }
                    }
                    #endregion

                    tarif1 = tarif1.Replace('.', ',');
                    tarif2 = tarif2.Replace('.', ',');
                    changer("tarif1", tarif1);                          //Первый тариф
                    changer("tarif2", tarif2);                          //Второй тариф

                    CreateAktClass.SaveDoc(textBox_path.Text, NumberCounter + tempstring);
                    CreateAktClass.CloseWordAll();
                }
                //CreateAktClass.CloseWord();
            }

        }

        /*private void SaveAndExit(string path, string filename)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            CreateAktClass.SaveDoc(path, filename);
            Object saveChanges = false;
            Object fileFormat = Word.WdSaveFormat.wdFormatDocument;
            Object routeDocument = Type.Missing;
            wordapp.Documents.Close(ref saveChanges, ref fileFormat, ref routeDocument);
            //(ref saveChanges, ref fileFormat, ref routeDocument);
        }*/


        /// <summary>
        /// Вставляем значения на позиции, указанные в шаблоне закладками
        /// </summary>
        /// <param name="bookmark">Имя закладки</param>
        /// <param name="value">Вставляемое значение</param>
        private void changer(string bookmark, string value)
        {
                CreateAktClass.ChangeBookmarks(bookmark, value);
        }

        /// <summary>
        /// Вставляет в акт поверки произвольные числа из класса точности 
        /// </summary>
        public void Randomizer()
        {
            for (int i = 1; i < 6; i++)
            {
                Random Rand = new Random();
                Double digit = Convert.ToDouble(Rand.Next(-19, 23))/Convert.ToDouble(100);
                if (digit * digit < 0.002543111)
                    i--;
                else
                    changer("rnd" + i, digit.ToString());
            }
        }

        /// <summary>
        /// Валидация тарифа для отображения в нормальной форме 5 нулей перед запятой
        /// </summary>
        /// <param name="Tar">Тариф</param>
        private void ChangeTar5(ref string Tar)
        {
            Tar = Tar.Replace('.', ',');
            Tar = Tar == "" ? "00000,0" : (Tar.Length <= 2 ? Tar + ",0" : Tar);
            Tar = Tar.Substring(Tar.Length - 2, 1).Contains(',') ? Tar :
                (Tar.Substring(Tar.Length - 1, 1).Contains(',') ? Tar + "0" : Tar + ",0");
            Tar = ("00000,0".Remove(7 - Tar.Length, Tar.Length)) + Tar;
            Tar = Tar == "" ? "0,0" : Tar.Replace('.', ',');
        }

        /// <summary>
        /// Валидация тарифа для отображения в нормальной форме 6 нулей перед запятой
        /// </summary>
        /// <param name="Tar"></param>
        private void ChangeTar6(ref string Tar)
        {
            Tar = Tar.Replace('.', ',');
            if (Tar == "")
                Tar = "000000,00";
            else if (!Tar.Contains(","))
                Tar = Tar + ",00";
            else if (Tar.Length - Tar.IndexOf(",")<3) {
                for (int j = Tar.Length - Tar.IndexOf(","); j < 3; j++)
                 Tar += "0"; 
            }

            if (Tar.Length < 9)
                for (int j = Tar.Length; j < 9; j++)
                    Tar = "0"+Tar;
        }

        /// <summary>
        /// Автоопределение квартала поверки по дате
        /// </summary>
        /// <param name="Poverka">Дата поверки формата DD.MM.YYYY</param>
        /// <returns></returns>
        private string PoverkaKvartal(string Poverka)
        { 
            String Kvart = "";
            int yandex = (Convert.ToInt32(Poverka.Substring(3, 2)) - 1) / 3 + 1;
            switch (yandex)
            {
                case 1: Kvart = "I"; break;
                case 2: Kvart = "II"; break;
                case 3: Kvart = "III"; break;
                case 4: Kvart = "IV"; break;
            }
            return Kvart;
        }

        /// <summary>
        /// Выбор папки для сохранения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result_box = folderBrowserDialog1.ShowDialog();
            if (result_box == System.Windows.Forms.DialogResult.OK)
                textBox_path.Text = folderBrowserDialog1.SelectedPath;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1_schetchikType.SelectedIndex = 0;
            Assembly ass = Assembly.GetExecutingAssembly();
            Version vers = ass.GetName().Version;

            label5.Text = vers.ToString();
            if (DebugClass.LoadData().Tables.Count > 0)
                dataGridView1.DataSource = DebugClass.LoadData().Tables[0];

            string[] loadstring = DebugClass.LoadForms();
            
            comboBox1_schetchikType.SelectedIndex = Convert.ToInt32(loadstring[0]);
            comboBox_maker.SelectedIndex = Convert.ToInt32(loadstring[1]);
            try
            {
                checkBox1.CheckState = (Convert.ToBoolean(loadstring[4])) ? CheckState.Checked : CheckState.Unchecked;
            }
            catch (Exception ex)
            {
                string test = ex.ToString();
            }

            ComboBoxInit();
            comboBox_control.SelectedIndex = Convert.ToInt32(loadstring[2]);
            
            textBox_datePoverka.Text = loadstring[3];
       }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (((DataTable)dataGridView1.DataSource).Rows.Count > 0 && tabControl1.SelectedIndex == 0)
                    switch (MessageBox.Show("Сохранить внесённые данные?", "Внимание", MessageBoxButtons.YesNo))
                    {
                        case DialogResult.Yes:
                            DataTable NDT = (DataTable)dataGridView1.DataSource;
                            NDT.TableName = "POVERKA";
                            SaveData(NDT);
                            DebugClass.insert_media(comboBox1_schetchikType.SelectedIndex.ToString(), comboBox_maker.SelectedIndex.ToString(), comboBox_control.SelectedIndex == -1 ? "0" : comboBox_control.SelectedIndex.ToString(), textBox_datePoverka.Text, checkBox1.Checked.ToString());
                            break;
                        case DialogResult.No:
                            DebugClass.SaveData();
                            break;
                    }
                StreamWriter strWriter = new StreamWriter(@"c:\shablon\data\path.txt");
                strWriter.Write(textBox_path.Text);
                strWriter.Close();
            }
            catch (Exception er) 
            {
                string exp = er.ToString();}
        }

        /// <summary>
        /// Автоматическое проставление квартала в столбце в зависимости от колонки поверка
        /// </summary>
        /// <param name="sender">Ячейка</param>
        /// <param name="e"></param>
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string tempdata;
                DataGridView DGW1 = ((DataGridView)sender);
                string colName = DGW1.Columns[DGW1.CurrentCell.ColumnIndex].Name;
                if (colName  == "Poverka")
                {
                    int rowInd = DGW1.CurrentCell.RowIndex;
                    tempdata = DGW1["Poverka", rowInd].Value.ToString();
                    
                    int yandex = (Convert.ToInt32(tempdata.Substring(3, 2)) - 1) / 3 + 1;
                    switch (yandex)
                    {
                        case 1: tempdata = "I"; break;
                        case 2: tempdata = "II"; break;
                        case 3: tempdata = "III"; break;
                        case 4: tempdata = "IV"; break;
                    }
                    DGW1["Kvartal", rowInd].Value = tempdata;
                }

            }
            catch { }
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            String[] Control = checkBox1.Checked ? Control1 : Control2;
            for (int i = 0; i < Control.Length; i++)
            {
                comboBox_control.Items.RemoveAt(0);
            }
            ComboBoxInit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable NDT = (DataTable)dataGridView2.DataSource;
            NDT.TableName = "SVIDETELSTVO";
            //DebugClass.SaveData(NDT);

            CreateAktClass.OpenDoc(@"c:\shablon\data","svid_2.doc");

            for (int i = 0; i < NDT.Rows.Count; i++)
            {
                //worddocument = CreateAktClass.worddocument;

                if (comboBox3.Text == "0,5")
                {
                   // worddocument.Tables[5].Rows[1].Delete();
                   // worddocument.Tables[7].Rows[1].Delete();
                }

                changer("type1", NDT.Rows[i]["TYPE"].ToString());
                changer("type2", NDT.Rows[i]["TYPE"].ToString());
                changer("number1", NDT.Rows[i]["ID"].ToString());
                changer("number2", NDT.Rows[i]["ID"].ToString());
                changer("temp1", temperature.ToString());
                changer("temp2", temperature.ToString());
                changer("humid1", humidity.ToString());
                changer("humid2", humidity.ToString());
                changer("press1", press.ToString());
                changer("press2", press.ToString());
                changer("date1", textBox_date2.Text);
                changer("date2", textBox_date2.Text);
                CreateAktClass.SaveDoc(@"c:\shablon\temp2_" + comboBox3.Text, "done");

                CreateAktClass.CloseDoc();

            }
            CreateAktClass.CloseWord();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result_box = folderBrowserDialog2.ShowDialog();
            if (result_box == System.Windows.Forms.DialogResult.OK)
                textBox1.Text = folderBrowserDialog2.SelectedPath;
        }


        public void SaveData(DataTable dt)
        {
            XmlTextWriter xmlw1 = new XmlTextWriter(@"c:\shablon\data\temp2.xml", UnicodeEncoding.UTF8);
            xmlw1.WriteStartDocument();
            xmlw1.WriteStartElement("Element");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                xmlw1.WriteStartElement("ROW");
                xmlw1.WriteAttributeString("ROW_ID", i.ToString());
                xmlw1.WriteAttributeString("ID", dt.Rows[i]["ID"].ToString());
                xmlw1.WriteAttributeString("Poverka", dt.Rows[i]["Poverka"].ToString() != "" ? dt.Rows[i]["Poverka"].ToString() : textBox_datePoverka.Text);
                xmlw1.WriteAttributeString("Tariff1", dt.Rows[i]["Tariff1"].ToString());
                xmlw1.WriteAttributeString("Tariff2", dt.Rows[i]["Tariff2"].ToString());
                xmlw1.WriteAttributeString("Kvartal", dt.Rows[i]["Kvartal"].ToString());
                xmlw1.WriteFullEndElement();
            }
            xmlw1.WriteFullEndElement();
            xmlw1.WriteEndDocument();
            xmlw1.Close();
        }

        private void dataGridView3_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                row = e.RowIndex;
                column = e.ColumnIndex;
                                
                //label15.Text = String.Format("строка {0}, столбец {1}, состояние строки {2}", row, column, dataGridView3.CurrentRow.IsNewRow.ToString());
            }
            catch {
            }
        }

        private void saveTable()
        { 
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (dataGridView3["PLACE_RES", row].Value == null || dataGridView3["PLACE_RES", row].Value.ToString() != e.ClickedItem.Text)
            {
                dataGridViewAddData(dataGridView3, "PLACE_RES", e.ClickedItem.Text);

                dataGridView3["PLACE_STATION", row].Value = "";
            }
        }

        private void contextMenuStrip2_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            contextMenuStrip2.Items.Clear();
            foreach (DataRow dr in DT_PODSTATIONS.Rows)
            {
                if (dr["RES"].ToString() == dataGridView3["PLACE_RES", row].Value.ToString())
                {
                    contextMenuStrip2.Items.Add(dr["PS"].ToString());
                }
            }
            contextMenuStrip2.Show(MousePosition.X, MousePosition.Y);
        }

        private void contextMenuStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "PLACE_STATION", e.ClickedItem.Text);
        }

        private void ContextYears_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ContextYears.Items.Clear();
            for (int j = DateTime.Now.Year; j >= DateTime.Now.Year-30; j--)
            {
                ContextYears.Items.Add(j.ToString());
            }
            ContextYears.Show(MousePosition.X, MousePosition.Y);
        }

        private void ContextYears_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "MANO_DATE", e.ClickedItem.Text);
        }

        private void списокПодстанцийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form Stat = new Stations();
            Stat.Activate();
            Stat.Show();
        }

        private void очиститьТаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedTab.Name)
            {
                case "TabPageSchetchiki": DT_SHETCHIK.Rows.Clear(); DT_SHETCHIK.AcceptChanges(); break;
                case "TabPageMano": MessageBox.Show("Манометры"); break;
                case "TabPageTT": MessageBox.Show("ТТ"); break;
                default: break;
            }
        }

        public void SaveData(string param, string values, string filename)
        {
            XmlTextWriter xmlw1 = new XmlTextWriter(pathmano+filename, UnicodeEncoding.UTF8);
            xmlw1.WriteStartDocument();
            xmlw1.WriteStartElement("Element");
                xmlw1.WriteStartElement("data");
                xmlw1.WriteAttributeString(param, values);
                xmlw1.WriteFullEndElement();
            xmlw1.WriteFullEndElement();
            xmlw1.WriteEndDocument();
            xmlw1.Close();
        }

        private void contextMenuStrip3_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "MANO_NAME", e.ClickedItem.Text);
        }


        /// <summary>
        /// Позволяет безопасно вставлять любой текст в ячейку, при этом если строка была новая, то грид рисует следующую
        /// </summary>
        /// <param name="dgv">Рабочий грид, к который вставляем значение</param>
        /// <param name="columndgv">Имя столбца</param>
        /// <param name="text">Значение ячейки</param>
        private void dataGridViewAddData(DataGridView dgv, string columndgv, string text)
        {
            dgv.Select();
            dgv.CurrentCell =
                dgv[columndgv, row];

            SendKeys.Send(text);
            dgv.EndEdit();
            SendKeys.Send("{ENTER}");
            SendKeys.Send("{UP}");
        }

        private void contextMenuStripDiapazon_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "SCALE", e.ClickedItem.Text);
        }

        private void contextMenuStripLength_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "MANO_LENGTH", e.ClickedItem.Text);
        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void dataGridView3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 46)
            {
                if (column>0)
                    if (dataGridView3.SelectedCells.Count > 0)
                        dataGridView3.SelectedCells[0].Value = (object)"";

                if (36 < e.KeyValue && e.KeyValue < 41)
                {
                    if (dataGridView3.SelectedCells.Count > 0)
                    {
                        column = dataGridView3.SelectedCells[0].ColumnIndex;
                        row = dataGridView3.SelectedCells[0].RowIndex;
                    }
                }
            }
        }


        private void insert_numbers()
        {
            if (dataGridView3.Rows.Count > 1)
            {
                lastid = 0;
                try
                {
                    XmlDocument xmdoc = new XmlDocument();
                    string full_pathxml = pathmano + String.Format("{0:yy}", DateTime.Now) + ".xml";
                    xmdoc.Load(full_pathxml);
                    XmlElement root = xmdoc.DocumentElement;
                    String last_ID = root.FirstChild.Attributes[0].Value;
                    lastid = Convert.ToInt32(last_ID);
                }
                catch
                { }
                for (int i = 1; i < dataGridView3.Rows.Count; i++)
                {
                    dataGridView3["ID_SVID", i - 1].Value = (object)(lastid+i).ToString();
                }
            }
        }

        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            insert_numbers();
        }

        private void dataGridView3_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            insert_numbers();
        }

        private void contextMenuStripNumbersNotExist_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            dataGridViewAddData(dataGridView3, "MANO_ID", e.ClickedItem.Text);
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            SetLastID();
            CreateAkt ctEmpty = new CreateAkt();
            ctEmpty.OpenDoc(@"C:\shablon\data\", "Empty.doc");
            
            if (dataGridView3.Rows.Count > 1)
            {
                bool isTKP = false;
                for (int i = 0; i < dataGridView3.Rows.Count - 2; i++)
                {
                    DataGridViewRow Dr = dataGridView3.Rows[i];
                    isTKP = Dr.Cells["MANO_NAME"].Value.ToString() == "ТКП-160Cr-М2";
                    filldoc(isTKP, true, Dr, ctEmpty);
                }
                DataGridViewRow Dr1 = dataGridView3.Rows[dataGridView3.Rows.Count-2];
                isTKP = Dr1.Cells["MANO_NAME"].Value.ToString() == "ТКП-160Cr-М2";
                filldoc(isTKP, false, Dr1, ctEmpty);
            }
            DateTime DTtemp = Convert.ToDateTime(textBox_manoSvidDate.Text);
            ctEmpty.SaveDoc(pathmano + String.Format("{0:yyyy}\\{0:MM}\\print", DTtemp), DateTime.Now.ToLongTimeString().Replace(":","_") + ".doc");
            MessageBox.Show("READY!");

            DataSet DST = new DataSet();
            try { DST.Tables.Remove("IDMANOMETR"); }
            catch{}
            DST.Tables.Add((DataTable)dataGridView3.DataSource);
            DST.WriteXml(String.Format(pathmano + "\\{0:yyyy}\\{0:MM}\\temp.xml", DTtemp));
        }
        
        /// <summary>
        /// заполнение сертификатов
        /// </summary>
        /// <param name="isTKP">Это ТКП? Если нет МТП</param>
        /// <param name="addNewPage">Добавлять переход на новую страницу?</param>
        /// <param name="Dr">Строка датагрида с данными</param>
        /// <param name="ctEmpty">Документ, куда вставляем полученный документ</param>
        private void filldoc(bool isTKP, bool addNewPage, DataGridViewRow Dr, CreateAkt ctEmpty)
        {
            int month = Convert.ToDateTime(textBox_manoSvidDate.Text).Month;
            CreateAkt ctTKP = new CreateAkt();
            if (isTKP)
            {
                ctTKP.OpenDoc(@"C:\shablon\data\", "TKP.doc");
                ctTKP.ChangeBookmarks("numbertkp1", Dr.Cells["MANO_ID"].Value.ToString());
                ctTKP.ChangeBookmarks("date1", String.Format("« {0:dd} » {1} {0:yyyy} г.", Convert.ToDateTime(textBox_manoSvidDate.Text), getMonthInRussian(month)));
                ctTKP.ChangeBookmarks("estimate1", String.Format("« {0:dd} » {1} {0:yyyy} г.", Convert.ToDateTime(textBox_manoSvidDate.Text).AddYears(3), getMonthInRussian(month)));
                ctTKP.ChangeBookmarks("sertnum1", String.Format("{0}/{1:yy}", Dr.Cells["ID_SVID"].Value.ToString(), DateTime.Now));
                ctTKP.ChangeBookmarks("year1", Dr.Cells["MANO_DATE"].Value.ToString());
                ctTKP.ChangeBookmarks("caliber1", comboBox1.Text);
                ctTKP.ChangeBookmarks("place1", String.Format("{0}, ПС {1}, {2}", Dr.Cells["PLACE_RES"].Value.ToString(), Dr.Cells["PLACE_STATION"].Value.ToString(), Dr.Cells["PLACE"].Value.ToString()));
            }
            else {
                ctTKP.OpenDoc(@"C:\shablon\data\", "mano1.doc");
                ctTKP.ChangeBookmarks("numbertkp1", Dr.Cells["MANO_ID"].Value.ToString());
                ctTKP.ChangeBookmarks("date1", String.Format("« {0:dd} » {1} {0:yyyy} г.", Convert.ToDateTime(textBox_manoSvidDate.Text), getMonthInRussian(month)));
                ctTKP.ChangeBookmarks("estimate1", String.Format("« {0:dd} » {1} {0:yyyy} г.", Convert.ToDateTime(textBox_manoSvidDate.Text).AddYears(3), getMonthInRussian(month)));
                ctTKP.ChangeBookmarks("sertnum1", String.Format("{0}/{1:yy}", Dr.Cells["ID_SVID"].Value.ToString(), DateTime.Now));
                ctTKP.ChangeBookmarks("year1", Dr.Cells["MANO_DATE"].Value.ToString());
                ctTKP.ChangeBookmarks("diapazon1", Dr.Cells["SCALE"].Value.ToString());
                ctTKP.ChangeBookmarks("caliber1", comboBox1.Text);
                ctTKP.ChangeBookmarks("place1", String.Format("{0}, ПС {1}, {2}", Dr.Cells["PLACE_RES"].Value.ToString(), Dr.Cells["PLACE_STATION"].Value.ToString(), Dr.Cells["PLACE"].Value.ToString()));
            }

            ctEmpty.InsertDoc(ctTKP.Word1.oDoc, addNewPage);
            Thread.Sleep(500);
            DateTime DTtemp = Convert.ToDateTime(textBox_manoSvidDate.Text);
            ctTKP.SaveDoc(pathmano + String.Format("{0:yyyy}\\{0:MM}", DTtemp), String.Format("Серт.{0} {1} {2}", Dr.Cells["ID_SVID"].Value.ToString(), isTKP?"ТКП":"Манометр", String.Format("{0} {1} {2}", Dr.Cells["PLACE_RES"].Value.ToString(), Dr.Cells["PLACE_STATION"].Value.ToString(), Dr.Cells["PLACE"].Value.ToString())));
            ctTKP.CloseDoc();
            ctTKP.CloseWord();
        }

        private void SetLastID()
        {
            insert_numbers();
            DateTime DTtemp = Convert.ToDateTime(textBox_manoSvidDate.Text);
            String fullPathManoSert = pathmano + String.Format("{0:yyyy}\\{0:MM}", DTtemp);
            if (!Directory.Exists(fullPathManoSert)) Directory.CreateDirectory(fullPathManoSert);
            if (!Directory.Exists(fullPathManoSert + @"\print")) Directory.CreateDirectory(fullPathManoSert + @"\print");
            int rcnt = dataGridView3.Rows.Count;
            if (rcnt > 1)
            {
                SaveData("LAST_ID", dataGridView3["ID_SVID", dataGridView3.Rows.Count - 2].Value.ToString(), String.Format("{0:yy}.xml", DTtemp));
            }
        }

        private string getMonthInRussian(int month)
        {
            if (month > 0 && month < 13)
                switch (month)
                {
                    case 1: return "января";
                    case 2: return "февраля";
                    case 3: return "марта";
                    case 4: return "апреля";
                    case 5: return "мая";
                    case 6: return "июня";
                    case 7: return "июля";
                    case 8: return "августа";
                    case 9: return "сентября";
                    case 10: return "октября";
                    case 11: return "ноября";
                    case 12: return "декабря";
                    default: return "Неверный месяц";
                }
            else return "Неверный месяц";
        }

        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form Empl = new Employee();
            Empl.Activate();
            Empl.Show();
        }
    }
}
