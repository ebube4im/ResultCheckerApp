using System.Web;
using System.Web.Mvc;

namespace GodlySeeds_Online_Result_Checker
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
