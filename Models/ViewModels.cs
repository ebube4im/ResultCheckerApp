using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace GodlySeeds_Online_Result_Checker.Models
{
    public class ViewModels
    {
    }


    public class PrintSubjectViewModel
    {
        public string Subjectname { get; set; }

        [Display(Name = "Assn")]
        public string CA1 { get; set; }

        [Display(Name = "1st Test")]
        public string CA2 { get; set; }

        [Display(Name = "2nd Test")]
        public string CA3 { get; set; }
      //  public string CA4 { get; set; }
        public string PROJECT { get; set; }
        public string EXAM { get; set; }

        [Display(Name = "Lowest in Class")]
        public string LOWESTINCLASS { get; set; }
        [Display(Name = "Highest in Class")]
        public string HIGHESTINCLASS { get; set; }

        [Display(Name = "Class Average")]
        public string CLASSAVERAGE { get; set; }
        public string TOTAL { get; set; }
        public string GRADE { get; set; }

    }

    public class PrintAnnualViewModel
    {
        [Display(Name = "Subject Name")]
        public string Subjectname { get; set; }

        [Display(Name = "1st Term Total")]
        public string FirstTermTotal { get; set; }

        [Display(Name = "1st Term Grade")]
        public string FirstTermGrade { get; set; }

        [Display(Name = "2nd Term Total")]
        public string secondTermTotal { get; set; }

        [Display(Name = "2nd Term Grade")]
        public string secondTermGrade { get; set; }

        [Display(Name = "3rd Term Total")]
        public string thirdTermTotal { get; set; }

        [Display(Name = "3rd Term Grade")]
        public string thirdTermGrade { get; set; }

         
        public string TOTAL { get; set; }
     
        [Display(Name = "TOTAL AVERAGE")]
        public string AVERAGE { get; set; }


        [Display(Name = "Class Average")]
        public string CLASSAVERAGE { get; set; }

      
         
        public string POSITION { get; set; }



        [Display(Name = "ANNUAL POSITION IN CLASS")]
        public string ANNUALPOSITION { get; set; }


        [Display(Name = "OVERALL POSITION COMBINED")]
        public string OVERALLPOSITION { get; set; }

        [Display(Name = "First Term Average")]
        public string firstTermTotalAverage { get; set; }
        
        [Display(Name = "Second Term Average")]
        public string secondTermTotalAverage { get; set; }

        [Display(Name = "Third Term Average")]
        public string thirdTermTotalAverage { get; set; }

        [Display(Name = "OVERALL AVERAGE")]
        public string Average { get; set; }
        public string Class { get; set; }
        public string PromotionStatus { get; set; }


    }


    public class StaffLoginViewModel
    {
        [Display(Name = "Name")]
        public string Username { get; set; }

        [Display(Name = "Password")]
        public string Password { get; set; }

        public string classList { get; set; }

        public string ErrorMessage { get; set; }

        public List<ClassListViewModel> classListViewModel { get; set; }

        public string[] CheckedList { get; set; }

    }

    public class CheckedVM {


        public string[] CheckedList { get; set; }


    }


    public class ClassListViewModel
    {
        [Display(Name = "Name")]
        public string StudentName { get; set; }

        [Display(Name = "Admission No")]
        public string AdmissionNo { get; set; }
        public string Gender { get; set; }


        public string Average { get; set; }
        public string Position { get; set; }

        public string NoInClass { get; set; }



    }


    public class PrintMultipleResultViewModel
    {
        public string ClassToPrint { get; set; }
        List<PrintResultViewModel> ListOfResultsInAClass { get; set; }
    }

        public class PrintResultViewModel
    {
        [Display(Name = "Name")]
        public string StudentName { get; set; }

        [Display(Name = "Admission No")]
        public string AdmissionNo { get; set; }
        public string Gender { get; set; }

        [Display(Name = "Number in Class")]
        public string NoinClass { get; set; }
        public string Class { get; set; }

        public List<PrintSubjectViewModel> mySubjects { get; set; }


        public List<PrintAnnualViewModel> annualSubjects { get; set; }
        
        public PrintAnnualViewModel printAnnualViewModel { get; set; }


        [Display(Name = "Total Marks Obtainable")]
        public string TotalMarksObtainable { get; set; }
        [Display(Name = "Total Marks Obtained")]
        public string TotalMarksObtained { get; set; }


        public string Average { get; set; }
        public string Position { get; set; }
        public string ErrorMessage { get; set; }

        public string Session { get; set; }

        public string Term { get; set; }

    }


    public class Admission
    {
        public string studentName { get; set; }


        [Required]
        [Display(Name ="Parent's Phone Number")]
        public string parentPhone { get; set; }
        public string Score { get; set; }
        public string ErrorMessage { get; set; }


    }

}