using System;
using DocumentoDel;
using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void MakeClass(){
        }

        [TestMethod]
        public void MakeEmployee()
        {
            Employee.Employees Empl = new Employee.Employees();
            Empl.HeIsMan = true;
            Empl.Birthday = Convert.ToDateTime("25.06.1985");
            Empl.Name = "Dmitry";
            Empl.LastName = "Lobanov";
            Empl.SurName = "Sergeevich";
            Empl.StartWorking = Convert.ToDateTime("01.09.2009");
            Assert.AreNotEqual(Empl, null);
        }
    }
}
