using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace IntegraDevProject.Web.Controllers
{
    public class HomeController : Controller
    {
        public static int ExportedFileCount { get; set; }
        public ActionResult Index()
        {
            return View();
        }
      [HttpGet]
        public JsonResult ExportToExcelFacilityList(/*FacilitySearchModel s*/)
        {
            try
            {
                List<string> allowed = new List<string>() { "FacilityName", "FacilityID", "Address", "City", "State",
                "Zip", "PhoneNumber"};
                

                DataTable dt = new DataTable
                {
                    Columns = {"FacilityName", "FacilityID", "Address", "City", "State", "Zip","PhoneNumber",
                                "EmailAddress", "Active", "ActiveDate", "LastUpdated", "LastUpdatedBy"}
                };

                ExportedFileCount++;
                string fileName = $"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\\Downloads\\facility{ExportedFileCount}.xlsx";

                DataTabletoExcel.ExporttabletoExcel(dt, fileName, allowed);

                return Json(new
                {
                    success = true,
                    FileName = fileName
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    success = false,
                    errors = "There is some issue in downloading the file. Please try later." + ex.Message
                }, JsonRequestBehavior.AllowGet); 
            }
        }

        //public ActionResult About()
        //{
        //    ViewBag.Message = "Your application description page.";

        //    return View();
        //}

        //public ActionResult Contact()
        //{
        //    ViewBag.Message = "Your contact page.";

        //   return View();
        //}
    }
}