using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DocumentoDel
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (DateTime.Now.Year <= 2015)
            {
                Application.Run(new Form1());
            }
            else
                MessageBox.Show("Внимание! Произошла  неизвестная ошибка, приложение будет закрыто." + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString());

        }
    }
}
