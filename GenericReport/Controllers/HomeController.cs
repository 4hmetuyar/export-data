using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GenerateReport;

namespace GenericReport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<Student> students = new List<Student>();
            for (int i = 0; i < 1048575; i++)
            {
                students.Add(new Student
                {
                    Email = "asdads",
                    Name = "asdasd",
                    Surname = "safasdfasf",
                    Email1 = "safddasfj adjflşdajsfşljsafdljdas",
                    Email2 = "asdfas dfdajsfladsjf ladfsj",
                    Email3 = "adsfj asdjf sahdofhsaodhfsadşfsad",
                    Email4 = "a lsakjdfşlsadfj sajdfsa",
                    Email5 = "a sadfjlsajdşfjsaldjfğoasf",
                    Email6 = " asjdfljsadfpjpamfoısadfmsafoa"
                });
            }



            ExcelExport<Student>.ToExcelFile(students, DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss"), null);
            return View();
        }

        public class Student
        {
            public string Name { get; set; }
            public string Surname { get; set; }
            public string Email { get; set; }
            public string Email1 { get; set; }
            public string Email2 { get; set; }
            public string Email3 { get; set; }
            public string Email4 { get; set; }
            public string Email5 { get; set; }
            public string Email6 { get; set; }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}