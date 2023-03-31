using GodlySeeds_Online_Result_Checker.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GodlySeeds_Online_Result_Checker.Controllers
{
    public class ELearningController : Controller
    {

        List<string> PaidFeesList = new List<string>();

        // GET: ELearning
        public ActionResult Index()
        {
            PrintResultViewModel mymodel = new PrintResultViewModel();

            if (mymodel.ErrorMessage != null)
            {
                ViewBag.Error = mymodel.ErrorMessage;
            }

            return View(mymodel);
        }

        [HttpPost]
        public ActionResult Index(PrintResultViewModel mymodel)
        {
            if (!ModelState.IsValid && mymodel == null)
            {
                mymodel.ErrorMessage = "Oops! A error occured with your submission. Try again";
                return View(mymodel);
            }

            // string myClass = "nothing";
            FileInfo myfile = new FileInfo(Server.MapPath("~/Content/Uploads/JSS1A.xlsx"));
            switch (mymodel.Class)
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
                    //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                    //return View("Index", mymodel);
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

                    var result2 = new PrintSubjectViewModel();

                    List<string> AdmissionNo = new List<string>();
                    List<PrintSubjectViewModel> mylist = new List<PrintSubjectViewModel>();
                    for (int i = 3; i <= noOfRow; i++)
                    {
                        AdmissionNo.Add(workSheet.Cells[i, 1].Value.ToString());
                    }



                    if (!AdmissionNo.Contains(mymodel.AdmissionNo.ToUpper()))
                    {
                        mymodel.ErrorMessage = string.Format("Oops! No Student with Admission Number  {0} Exists in this Class. Please Check the spelling and try again", mymodel.AdmissionNo.ToUpper());
                        return View("Index", mymodel);
                    }

                    UpdateFees();

                    if (!PaidFeesList.Contains(mymodel.AdmissionNo.ToUpper()))
                    {
                        mymodel.ErrorMessage = string.Format("Oops! Student with Admission Number {0} is yet to complete fees for the Second Term - 2019/2020", mymodel.AdmissionNo.ToUpper());
                        return View("Index", mymodel);
                    }


                    int number = AdmissionNo.IndexOf(mymodel.AdmissionNo.ToUpper()) + 3;

                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();



                }
                 

                using (System.IO.StreamWriter file =
                         new System.IO.StreamWriter(Server.MapPath("~/Content/Uploads/Elearning.txt"), true))
                {
                    file.WriteLine("\nStudent:  " + mymodel.StudentName + " " + mymodel.AdmissionNo + " | " + "Class: " + mymodel.Class + " | Time Checked: " + DateTime.UtcNow);
                }
            }

            return RedirectToAction("Dashboard", mymodel);


        }




        public void UpdateFees()
        {
            FileInfo FeesFile = new FileInfo(Server.MapPath("~/Content/Uploads/FeesList.xlsx"));

            if (FeesFile != null)
            {
                using (var package = new ExcelPackage(FeesFile))
                {
                    var currentSheet = package.Workbook.Worksheets;
                    var workSheet = currentSheet.First();
                    var noOfCol = workSheet.Dimension.End.Column;
                    var noOfRow = workSheet.Dimension.End.Row;

                    for (int i = 2; i <= noOfRow; i++)
                    {
                        PaidFeesList.Add(workSheet.Cells[i, 4].Value.ToString());

                    }

                }

            }
        }



       
        public ActionResult Dashboard(PrintResultViewModel mymodel)
        {
            if(mymodel.AdmissionNo == null)
            {
                return RedirectToAction("Index");
            }
            switch (mymodel.Class)
            {
                case "JSS 1A":
                    mymodel.Class = "JSS1";
                    break;
                case "JSS 1B":
                    mymodel.Class = "JSS1";
                    break;
                case "JSS 2A":
                    mymodel.Class = "JSS2";
                    break;
                case "JSS 2B":
                    mymodel.Class = "JSS2";
                    break;
                case "JSS 3A":
                    mymodel.Class = "JSS3";
                    break;
                case "JSS 3B":
                    mymodel.Class = "JSS3";
                    break;
                case "SSS 1A":
                    mymodel.Class = "SS1";
                    break;

                default:
                    // To DO
                    break;
            }




            return View(mymodel);
        }

    }
}
