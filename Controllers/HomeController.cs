using GodlySeeds_Online_Result_Checker.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using Microsoft.Ajax.Utilities;

namespace GodlySeeds_Online_Result_Checker.Controllers
{
    public class HomeController : Controller
    {
        List<string> PaidFeesList = new List<string>();

        public ActionResult Index(PrintResultViewModel mymodel)
        {


            if (mymodel.ErrorMessage != null)
            {
                ViewBag.Error = mymodel.ErrorMessage;
            }
            ViewBag.Title = "Home";
            return View(mymodel);
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

        [HttpGet]
        public ActionResult Print()
        {
            ViewBag.Title = "Print Result";
            PrintResultViewModel mymodel = new PrintResultViewModel();
            return View("Index", mymodel);
        }

        [HttpPost]
        public ActionResult Print(PrintResultViewModel mymodel)
        {

           

            ViewBag.Title = "Print Result";
            if (!ModelState.IsValid && mymodel == null)
            {
                mymodel.ErrorMessage = "Oops! A error occured with your submission. Try again";
                return View("Index", mymodel);
            }
            
            // string myClass = "nothing";
            FileInfo myfile = new FileInfo("Nothing");
            // = new FileInfo(Server.MapPath("~/Content/Uploads/JSS1A.xlsx"));


            try
            {
                if (mymodel.Session == "2020")
                {
                    if (mymodel.Term == "1st Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/1stTerm/SSS1A.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                    if (mymodel.Term == "2nd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/2ndTerm/SSS1A.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                    if (mymodel.Term == "3rd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                //          myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                //        myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/3rdTerm/SSS1A.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }
                    if (mymodel.Term == "Annual")
                    {

              
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2020/Annual/SSS1A.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                }

                if (mymodel.Session == "2021")
                {
                    if (mymodel.Term == "1st Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/SSS1A.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/1stTerm/SSS2Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                    if (mymodel.Term == "2nd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/JSS3B.xlsx"));
                                break;
                            case "SSS1Arts":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/SSS1Arts.xlsx"));
                                break;
                            case "SSS1Science":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/SSS1Science.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/2ndTerm/SSS2Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                    if (mymodel.Term == "3rd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/JSS3B.xlsx"));
                                break;
                            case "SSS1Arts":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/SSS1Arts.xlsx"));
                                break;
                            case "SSS1Science":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/SSS1Science.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/3rdTerm/SSS2Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }

                        if (mymodel.Term == "Annual")
                        {

                            //if (mymodel.Class == "SSS 1A")
                            //{

                            //    mymodel.ErrorMessage = string.Format("Oops! No Result has been uploaded for SS1 Annual. Please Check you individual Results for First, Second and Third Term. Do check later for annual. We apologise for the inconvenience.  ");
                            //    return View("Index", mymodel);
                            //}


                            switch (mymodel.Class)
                            {
                                case "JSS 1A":
                                    //     myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS1A.xlsx"));
                                    break;
                                case "JSS 1B":
                                    //    myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS1B.xlsx"));
                                    break;
                                case "JSS 2A":
                                    //    myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS2A.xlsx"));
                                    break;
                                case "JSS 2B":
                                    //     myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS2B.xlsx"));
                                    break;
                                case "JSS 3A":
                                    // myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS3A.xlsx"));
                                    break;
                                case "JSS 3B":
                                    //myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS3B.xlsx"));
                                    break;
                                case "SSS1Arts":
                                    //    myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS1Arts.xlsx"));
                                    break;
                                case "SSS1Science":
                                    // myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS1Science.xlsx"));
                                    break;
                                case "SSS2Arts":
                                    // myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS2Arts.xlsx"));
                                    break;
                                case "SSS2Science":
                                    //   myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS2Science.xlsx"));
                                    break;

                                default:
                                    // To DO
                                    break;
                            }
                        }
                    }

                    if (mymodel.Term == "Annual")
                    {


                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS1AAnnual.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS1BAnnual.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS2AAnnual.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS2BAnnual.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS3AAnnual.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/JSS3BAnnual.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS1AAnnual.xlsx"));
                                break;
                            case "SSS2Arts":
                                 myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS2ArtsAnnual.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2021/Annual/SSS2ScienceAnnual.xlsx"));
                                break;
                            default:
                                // To DO
                                break;
                        }
                    }

                }

                if (mymodel.Session == "2022")
                {
                    if (mymodel.Term == "1st Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS1A.xlsx"));
                                break;
                            case "SSS1Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS1Arts.xlsx"));
                                break;
                            case "SSS1Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS1Science.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS2Science.xlsx"));
                                break;
                            case "SSS3Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS3Arts.xlsx"));
                                break;
                            case "SSS3Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/1stTerm/SSS3Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }


                    if (mymodel.Term == "2nd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS1A.xlsx"));
                                break;
                            case "SSS1Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS1Arts.xlsx"));
                                break;
                            case "SSS1Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS1Science.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS2Science.xlsx"));
                                break;
                            case "SSS3Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS3Arts.xlsx"));
                                break;
                            case "SSS3Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/2ndTerm/SSS3Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }


                    if (mymodel.Term == "3rd Term")
                    {
                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS1A.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS1B.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS2A.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS2B.xlsx"));
                                break;
                            case "JSS 3A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS3A.xlsx"));
                                break;
                            case "JSS 3B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/JSS3B.xlsx"));
                                break;
                            case "SSS 1A":
                                //mymodel.ErrorMessage = "Oops! No result has been uploaded yet for the selected class. Please check back later";
                                //return View("Index", mymodel);
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS1A.xlsx"));
                                break;
                            case "SSS1Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS1Arts.xlsx"));
                                break;
                            case "SSS1Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS1Science.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS2Arts.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS2Science.xlsx"));
                                break;
                            case "SSS3Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS3Arts.xlsx"));
                                break;
                            case "SSS3Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/3rdTerm/SSS3Science.xlsx"));
                                break;

                            default:
                                // To DO
                                break;
                        }
                    }

                    if (mymodel.Term == "Annual")
                    {


                        switch (mymodel.Class)
                        {
                            case "JSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/JSS1AAnnual.xlsx"));
                                break;
                            case "JSS 1B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/JSS1BAnnual.xlsx"));
                                break;
                            case "JSS 2A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/JSS2AAnnual.xlsx"));
                                break;
                            case "JSS 2B":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/JSS2BAnnual.xlsx"));
                                break;
                            case "SSS 1A":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/SSS1AAnnual.xlsx"));
                                break;
                            case "SSS1Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/SSS1ArtsAnnual.xlsx"));
                                break;
                            case "SSS1Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/SSS1ScienceAnnual.xlsx"));
                                break;
                            case "SSS2Arts":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/SSS2ArtsAnnual.xlsx"));
                                break;
                            case "SSS2Science":
                                myfile = new FileInfo(Server.MapPath("~/Content/Uploads/2022/Annual/SSS2ScienceAnnual.xlsx"));
                                break;
                            default:
                                // To DO
                                break;
                        }
                    }

                }

                if (myfile.Exists)
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
                        List<PrintAnnualViewModel> Annuallist = new List<PrintAnnualViewModel>();



                        if (mymodel.Term == "Annual")
                        {


                            for (int i = 3; i <= noOfRow; i++)
                            {
                                AdmissionNo.Add(workSheet.Cells[i, 2].Value.ToString());
                            }
                        }
                        else
                        {
                            for (int i = 3; i <= noOfRow; i++)
                            {
                                AdmissionNo.Add(workSheet.Cells[i, 1].Value.ToString());
                            }
                        }


                        if (!AdmissionNo.Contains(mymodel.AdmissionNo.ToUpper()))
                        {
                            mymodel.ErrorMessage = string.Format("Oops! No Student with Admission Number  {0} Exists in this Class. Please Check the spelling and try again", mymodel.AdmissionNo.ToUpper());
                            return View("Index", mymodel);
                        }

                        //   UpdateFees();

                        //if (!PaidFeesList.Contains(mymodel.AdmissionNo.ToUpper()))
                        //{
                        //    mymodel.ErrorMessage = string.Format("Oops! Student with Admission Number {0} is yet to complete fees for the Second Term - 2019/2020", mymodel.AdmissionNo.ToUpper());
                        //    return View("Index", mymodel);
                        //}


                        int number = AdmissionNo.IndexOf(mymodel.AdmissionNo.ToUpper()) + 3;

                         
                            if (mymodel.Session == "2020")
                            {

                            if (mymodel.Term == "1st Term")
                            {

                            }

                            if (mymodel.Term == "2nd Term")
                            {

                            }

                            if (mymodel.Term == "3rd Term")
                            {

                            }

 
                            for (int i = 3; i <= 150; i++)
                                {
                                    var result = new PrintSubjectViewModel();
                                    result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                    result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                    i = i + 1;
                                    result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                    i = i + 1;
                                    result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                    i = i + 1;
                                    result.PROJECT = workSheet.Cells[number, i].Value.ToString();
                                    i = i + 1;
                                    result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                    i = i + 1;
                                    result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                    i = i + 1;
                                    result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                    i = i + 1;

                                    result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                    i = i + 1;
                                    result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                    i = i + 1;
                                    result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                    mylist.Add(result);
                                }

                                // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                mymodel.Gender = workSheet.Cells[number, 158].Value.ToString();
                                mymodel.Position = workSheet.Cells[number, 157].Value.ToString();
                                mymodel.TotalMarksObtainable = workSheet.Cells[number, 154].Value.ToString();
                                mymodel.TotalMarksObtained = workSheet.Cells[number, 155].Value.ToString();
                                mymodel.Average = workSheet.Cells[number, 156].Value.ToString();
                                mymodel.Class = workSheet.Cells[2, 2].Value.ToString();
                                mymodel.mySubjects = mylist;

                            }

                            if (mymodel.Session == "2021")
                        {

                            if (mymodel.Term == "1st Term")
                            {
                                if (mymodel.Class == "SSS1A")
                                {
                                    for (int i = 3; i <= 120; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.PROJECT = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 128].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 127].Value.ToString();
                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 124].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 125].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 126].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();
                                    mymodel.mySubjects = mylist;

                                }
                                else
                                if (mymodel.Class == "JSS3A" || mymodel.Class == "JSS3B")
                                {
                                    for (int i = 3; i <= 140; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.PROJECT = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 128].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 127].Value.ToString();
                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 124].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 125].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 126].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();
                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS2Science")
                                {

                                    for (int i = 3; i <= 80; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.PROJECT = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 84].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 85].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 86].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 87].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 88].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 130; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.PROJECT = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 138].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 137].Value.ToString();
                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 134].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 135].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 136].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();
                                    mymodel.mySubjects = mylist;

                                }


                            }

                            if (mymodel.Term == "2nd Term")
                            {
                                if (mymodel.Class == "SSS2Arts" || mymodel.Class == "SSS1Arts" || mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 3; i <= 81; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 85].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 86].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 87].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 88].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 89].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS2Science")
                                {

                                    for (int i = 3; i <= 72; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 76].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 77].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 78].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 79].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 80].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 126; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 131].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 132].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 133].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 134].Value.ToString();


                                    mymodel.mySubjects = mylist;



                                }

                            }

                            if (mymodel.Term == "3rd Term")
                            {
                                if (mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 3; i <= 81; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 85].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 86].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 87].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 88].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 89].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS1Arts" || mymodel.Class == "SSS2Arts")
                                {
                                    for (int i = 3; i <= 90; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 94].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 95].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 96].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 97].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 98].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }


                                else if (mymodel.Class == "SSS2Science")
                                {

                                    for (int i = 3; i <= 72; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 76].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 77].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 78].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 79].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 80].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 126; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.AdmissionNo = workSheet.Cells[number, 1].Value.ToString();

                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 131].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 132].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 133].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 134].Value.ToString();


                                    mymodel.mySubjects = mylist;



                                }

                            }

                            if (mymodel.Term == "Annual")
                            {

                                //i Cannot obtain value of the local variable or argument because it is not available at this instruction pointer, possibly because it has been optimized away.	int
                                if (mymodel.Class == "SSS 1A")
                                {
                                    for (int i = 4; i <= 133; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 140].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3900";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 134].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 139].Value.ToString(),
                                      //  OVERALLPOSITION = workSheet.Cells[number, 140].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 135].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 136].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 137].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 138].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 142].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();
                                    mymodel.annualSubjects = Annuallist;

                                }
                                else if (mymodel.Class == "SSS2Science")
                                {
                                    for (int i = 4; i <= 83; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 90].Value.ToString();
                                    mymodel.TotalMarksObtainable = "2400";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 84].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        // ANNUALPOSITION = workSheet.Cells[number, 89].Value.ToString(),
                                        //    OVERALLPOSITION = workSheet.Cells[number, 150].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 85].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 86].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 87].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 88].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 92].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();

                                    mymodel.annualSubjects = Annuallist;

                                }
                                else if (mymodel.Class == "SSS2Arts")
                                {
                                    for (int i = 4; i <= 103; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 110].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3000";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 104].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        //  ANNUALPOSITION = workSheet.Cells[number, 109].Value.ToString(),
                                        // OVERALLPOSITION = workSheet.Cells[number, 150].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 105].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 106].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 107].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 108].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 112].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();

                                    mymodel.annualSubjects = Annuallist;

                                }
                                else
                                {


                                    for (int i = 4; i <= 143; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 151].Value.ToString();
                                    mymodel.TotalMarksObtainable = "4200";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 144].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 149].Value.ToString(),
                                        OVERALLPOSITION = workSheet.Cells[number, 150].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 145].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 146].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 147].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 148].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 153].Value.ToString(),

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();


                                    mymodel.annualSubjects = Annuallist;

                                }
                            }

                        }


                        if (mymodel.Session == "2022")
                        {

                            if (mymodel.Term == "1st Term")
                            {
                                if (mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 3; i <= 81; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 85].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 86].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 87].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 88].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 89].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS1Arts" || mymodel.Class == "SSS2Arts")
                                {
                                    for (int i = 3; i <= 101; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number,115].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 116].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 117].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 118].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 119].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }


                                else if (mymodel.Class == "SSS2Science")
                                {

                                    for (int i = 3; i <= 72; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 76].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 77].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 78].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 79].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 80].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 135; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.AdmissionNo = workSheet.Cells[number, 1].Value.ToString();

                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 131].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 132].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 133].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 134].Value.ToString();


                                    mymodel.mySubjects = mylist;



                                }

                            }

                            if (mymodel.Term == "2nd Term")
                            {
                                if (mymodel.Class == "SSS2Science" || mymodel.Class == "SSS2Arts" || mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 3; i <= 108; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 112].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 113].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 114].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 115].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 116].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS1Arts")
                                {
                                    for (int i = 3; i <= 99; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 103].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 104].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 105].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 106].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 107].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }


                                else if (mymodel.Class == "SSS3Science")
                                {

                                    for (int i = 3; i <= 72; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 76].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 77].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 78].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 79].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 80].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 135; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.AdmissionNo = workSheet.Cells[number, 1].Value.ToString();

                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 139].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 140].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 141].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 142].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 143].Value.ToString();


                                    mymodel.mySubjects = mylist;



                                }

                            }

                            if (mymodel.Term == "3rd Term")
                            {
                                if (mymodel.Class == "SSS2Science" || mymodel.Class == "SSS2Arts" || mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 3; i <= 108; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 112].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 113].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 114].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 115].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 116].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }

                                else if (mymodel.Class == "SSS1Arts")
                                {
                                    for (int i = 3; i <= 99; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 103].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 104].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 105].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 106].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 107].Value.ToString();

                                    mymodel.mySubjects = mylist;

                                }


                                else if (mymodel.Class == "SSS3Science")
                                {

                                    for (int i = 3; i <= 72; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }


                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 76].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 77].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 78].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 79].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 80].Value.ToString();

                                    mymodel.mySubjects = mylist;


                                }

                                else
                                {

                                    for (int i = 3; i <= 135; i++)
                                    {
                                        var result = new PrintSubjectViewModel();
                                        result.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        result.CA1 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA2 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.CA3 = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        result.EXAM = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.LOWESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.HIGHESTINCLASS = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;

                                        result.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();
                                        i = i + 1;
                                        result.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        result.GRADE = workSheet.Cells[number, i].Value.ToString();
                                        mylist.Add(result);
                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 2].Value.ToString();
                                    mymodel.AdmissionNo = workSheet.Cells[number, 1].Value.ToString();

                                    mymodel.Class = workSheet.Cells[2, 2].Value.ToString();

                                    mymodel.TotalMarksObtainable = workSheet.Cells[number, 139].Value.ToString();
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 140].Value.ToString();
                                    mymodel.Average = workSheet.Cells[number, 141].Value.ToString();
                                    mymodel.Position = workSheet.Cells[number, 142].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 143].Value.ToString();


                                    mymodel.mySubjects = mylist;



                                }

                            }

                            if (mymodel.Term == "Annual")
                            {

                                //i Cannot obtain value of the local variable or argument because it is not available at this instruction pointer, possibly because it has been optimized away.	int
                                if (mymodel.Class == "SSS1Arts")
                                {
                                    for (int i = 4; i <= 113; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 120].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3000";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 114].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 119].Value.ToString(),
                                        //  OVERALLPOSITION = workSheet.Cells[number, 140].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 115].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 116].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 117].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 118].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 122].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();
                                    mymodel.annualSubjects = Annuallist;

                                }

                                else if (mymodel.Class == "SSS1Science")
                                {
                                    for (int i = 4; i <= 123; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3300";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 124].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 129].Value.ToString(),
                                        //  OVERALLPOSITION = workSheet.Cells[number, 140].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 125].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 126].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 127].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 128].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 132].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();
                                    mymodel.annualSubjects = Annuallist;

                                }

                                else if (mymodel.Class == "SSS2Arts")
                                {
                                    for (int i = 4; i <= 123; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3300";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 124].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 129].Value.ToString(),
                                        //  OVERALLPOSITION = workSheet.Cells[number, 140].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 125].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 126].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 127].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 128].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 132].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();
                                    mymodel.annualSubjects = Annuallist;

                                }

                                else if (mymodel.Class == "SSS2Science")
                                {
                                    for (int i = 4; i <= 123; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 130].Value.ToString();
                                    mymodel.TotalMarksObtainable = "3300";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 124].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 129].Value.ToString(),
                                        //  OVERALLPOSITION = workSheet.Cells[number, 140].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 125].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 126].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 127].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 128].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 132].Value.ToString()

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();
                                    mymodel.annualSubjects = Annuallist;

                                }

                                else
                                {


                                    for (int i = 4; i <= 153; i++)
                                    {

                                        var annualResult = new PrintAnnualViewModel();


                                        annualResult.Subjectname = workSheet.Cells[1, i].Value.ToString();
                                        annualResult.FirstTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.FirstTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.secondTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.secondTermGrade = workSheet.Cells[number, i].Value.ToString();

                                        i = i + 1;
                                        annualResult.thirdTermTotal = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;
                                        annualResult.thirdTermGrade = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.TOTAL = workSheet.Cells[number, i].Value.ToString();
                                        i = i + 1;

                                        annualResult.AVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.CLASSAVERAGE = (decimal.ToUInt16(decimal.Parse(workSheet.Cells[number, i].Value.ToString()))).ToString();

                                        i = i + 1;

                                        annualResult.POSITION = workSheet.Cells[number, i].Value.ToString();


                                        Annuallist.Add(annualResult);

                                    }

                                    // result2.Subjectname = workSheet.Cells["M1"].Value.ToString();
                                    mymodel.NoinClass = workSheet.Cells[1, 1].Value.ToString();
                                    mymodel.StudentName = workSheet.Cells[number, 3].Value.ToString();
                                    mymodel.Gender = workSheet.Cells[number, 161].Value.ToString();
                                    mymodel.TotalMarksObtainable = "4500";
                                    mymodel.TotalMarksObtained = workSheet.Cells[number, 154].Value.ToString();

                                    mymodel.printAnnualViewModel = new PrintAnnualViewModel()
                                    {
                                        ANNUALPOSITION = workSheet.Cells[number, 159].Value.ToString(),
                                        OVERALLPOSITION = workSheet.Cells[number, 160].Value.ToString(),
                                        Class = mymodel.Class,
                                        firstTermTotalAverage = decimal.Parse(workSheet.Cells[number, 155].Value.ToString()).ToString(),
                                        secondTermTotalAverage = decimal.Parse(workSheet.Cells[number, 156].Value.ToString()).ToString(),
                                        thirdTermTotalAverage = decimal.Parse(workSheet.Cells[number, 157].Value.ToString()).ToString(),
                                        Average = workSheet.Cells[number, 158].Value.ToString(),
                                        PromotionStatus = workSheet.Cells[number, 163].Value.ToString(),

                                    };

                                    mymodel.Class = workSheet.Cells[1, 3].Value.ToString();


                                    mymodel.annualSubjects = Annuallist;

                                }
                            }

                        }

                    }
 
                    using (System.IO.StreamWriter file =
                     new System.IO.StreamWriter(Server.MapPath("~/Content/Uploads/Checked.txt"), true))
                    {
                        file.WriteLine("\nStudent:  " + mymodel.StudentName + " " + mymodel.AdmissionNo + " | " + "Class: " + mymodel.Class + " | TERM: " + mymodel.Term + " | Time Checked: " + DateTime.UtcNow);

                    }


                }
                 
                else
                {
                    mymodel.ErrorMessage = string.Format("Oops! No Results have been Uploaded for your current selection");
                    return View("Index", mymodel);
                }
                 

            }
            catch (Exception e)
            {
                
                mymodel.ErrorMessage = string.Format("Oops! An error occurred with your current operation. Please try again or Contact Admin",e);
                return View("Index", mymodel);
            }


            return View(mymodel);
             

        }

         
        public ActionResult Test()
        {

            return View();
        }
    }
}


