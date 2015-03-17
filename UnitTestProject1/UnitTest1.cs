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
            Empl.Birthday = Convert.ToDateTime("22.01.1975");
            Empl.Name = "Denis";
            Empl.LastName = "Lobanov";
            Empl.SurName = "Evgenjevich";
            Empl.StartWorking = Convert.ToDateTime("01.09.2009");
            Assert.AreNotEqual(Empl, null);
        }
    }
}
