using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApplication3.Service;

namespace TimeTable.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public FileContentResult UploadFile(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0 
                && (Path.GetFileName(file.FileName).EndsWith(".xls") 
                || Path.GetFileName(file.FileName).EndsWith(".xlsx")))
                try
                {
                    string filePath = Path.Combine(Server.MapPath("~/FileUpload"), Path.GetFileName(file.FileName));
                    file.SaveAs(filePath);
                    if (filePath.EndsWith(".xls"))
                    {
                        var book = new Aspose.Cells.Workbook(filePath);
                        // save XLS as XLSX
                        filePath = filePath.Replace(".xls", $"_{DateTimeOffset.UtcNow.AddHours(7).ToUnixTimeMilliseconds()}.xlsx");
                        book.Save(filePath, Aspose.Cells.SaveFormat.Auto);

                    }
                    var model = FileServices.Instance.ImportFileExcel(filePath);
                    string[] header =
                    {
                            "Week",
                            "Monday",
                            "Tuesday",
                            "Wednesday",
                            "Thursday",
                            "Friday",
                            "Saturday",
                            "Sunday"
                        };
                    byte[] fileByte = FileServices.Instance.ExportFileExcel(model, header);
                    FileContentResult fileContentResult = File(fileByte, "application/vnd.ms-excel", "TimeTable.xlsx");
                    return fileContentResult;

                }
                catch (Exception ex)
                {
                }
            else
            {
            }
            return default;
        }

    }
}