using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using GodlySeeds_Online_Result_Checker.Models;
using OfficeOpenXml;

namespace GodlySeeds_Online_Result_Checker.Controllers
{
    public class StaffController : Controller
    {
        // GET: Staff
        public ActionResult Index()
        {
            StaffLoginViewModel mymodel = new StaffLoginViewModel();
            return View(mymodel);
        }

        [HttpGet]
        public ActionResult Checked()
        {
            string[] lines = System.IO.File.ReadAllLines(Server.MapPath("~/Content/Uploads/Checked.txt"));
            CheckedVM checklist = new CheckedVM();
            checklist.CheckedList = lines;

            return View(checklist);
        }


        [HttpGet]
        public ActionResult ViewClassList()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ViewClassList(StaffLoginViewModel mymodel)
        {

            if (!ModelState.IsValid && mymodel == null)
            {
                mymodel.ErrorMessage = "Oops! A error occured with your submission. Try again";
                return View("Index", mymodel);
            }


            if (mymodel.Username == "GSCCADMIN" && mymodel.Password == "admin@2020")
            {

                // string myClass = "nothing";
                FileInfo myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS1A.xlsx"));
                switch (mymodel.classList)
                {
                    case "JSS 1A":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS1A.xlsx"));
                        break;
                    case "JSS 1B":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS1B.xlsx"));
                        break;
                    case "JSS 2A":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS2A.xlsx"));
                        break;
                    case "JSS 2B":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS2B.xlsx"));
                        break;
                    case "JSS 3A":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS3A.xlsx"));
                        break;
                    case "JSS 3B":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS3B.xlsx"));
                        break;
                    case "SSS 1A":
                        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/SS1A.xlsx"));
                        break;

                    default:
                        // To DO
                        break;
                }


                if (myfile != null)
                {
                    using (var package = new ExcelPackage(myfile))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;

                        List<ClassListViewModel> ClassList = new List<ClassListViewModel>();

                      
 
                        for (int i = 3; i <= noOfRow; i++)
                        {
                            var result = new ClassListViewModel();
                            result.NoInClass = workSheet.Cells[1, 1].Value.ToString();
                            result.StudentName = workSheet.Cells[i, 2].Value.ToString();
                            result.AdmissionNo = workSheet.Cells[i, 1].Value.ToString();
                            result.Gender = workSheet.Cells[i, 158].Value.ToString();
                            result.Position = workSheet.Cells[i, 157].Value.ToString();

                            ClassList.Add(result);
                             
                        }

                        mymodel.classListViewModel = ClassList;

                        string[] lines = System.IO.File.ReadAllLines(Server.MapPath("~/Content/Uploads/Checked.txt"));

                        mymodel.CheckedList = lines;
                        return View(mymodel);
                    }

            
                }

            }
            return View();
        }

    }
    
}