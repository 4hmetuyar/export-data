using GenerateReport;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace GenericReport.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            //Task.Factory.StartNew(() => ExportData());
            var filePath = System.Web.HttpContext.Current.Server.MapPath($"/content/upload/{DateTime.Now:yyyy-dd-M-HH-mm-ss}.xlsx");

            Task.Run(() => ExportData(filePath));

            return View();
        }

        private void ExportData(string filePath)
        {
            //example data.
            List<Student> students = new List<Student>();
            for (int i = 0; i < 100575; i++)
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
            ExcelExport<Student>.ToExcelFile(students, filePath, sheetName: "example");
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
    }
}