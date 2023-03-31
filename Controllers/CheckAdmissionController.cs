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
    public class CheckAdmissionController : Controller
    {
        // GET: CheckAdmission
        public ActionResult Index()
        {
            
            return View(new Admission());
        }

        [HttpPost]
        public ActionResult Index(Admission model)
        {
            if (!ModelState.IsValid && model == null)
            {
                model.ErrorMessage = "Oops! A error occured with your submission. Try again";
                return View(model);
            }

            // string myClass = "nothing";
            FileInfo myfile = new FileInfo(Server.MapPath("~/Content/Uploads/AdmissionList.xlsx"));
        

            if (myfile != null)
            {
                using (var package = new ExcelPackage(myfile))
                {
                    var currentSheet = package.Workbook.Worksheets;
                    var workSheet = currentSheet.First();
                    var noOfCol = workSheet.Dimension.End.Column;
                    var noOfRow = workSheet.Dimension.End.Row;

                    Admission AdmissionList = new Admission();

                    List<string> AdmissionNo = new List<string>();
                    List<Admission> mylist = new List<Admission>();
                    for (int i = 2; i <= noOfRow; i++)
                    {
                        AdmissionNo.Add("0" + workSheet.Cells[i, 1].Value.ToString());
                    }



                    if (!AdmissionNo.Contains(model.parentPhone))
                    {
                        model.ErrorMessage = string.Format("Oops! No Student with Admission Number  {0} Exists in this Class. Please Check the spelling and try again", model.parentPhone);
                        return View("Index", model);
                    }

                     

                    int number = AdmissionNo.IndexOf(model.parentPhone) + 2;

                    model.studentName = workSheet.Cells[number, 2].Value.ToString();
                    model.parentPhone = workSheet.Cells[number, 1].Value.ToString();
                    model.Score = Int32.Parse(workSheet.Cells[number, 3].Value.ToString()).ToString();



                }



                using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(Server.MapPath("~/Content/Uploads/AdmissionList.txt"), true))
                {
                    file.WriteLine("\nStudent:  " + model.studentName + " | 0" + model.parentPhone + "  | Time Checked: " + DateTime.UtcNow);
                }
            }

            return RedirectToAction("Download", model);
 
        }

        public ActionResult Download(Admission model)
        {


            return View(model);
        }
         
    }
}
