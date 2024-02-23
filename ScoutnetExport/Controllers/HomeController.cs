using Microsoft.AspNetCore.Mvc;
using ScoutnetExport.Helpers;
using ScoutnetExport.Models;
using System.Diagnostics;

namespace ScoutnetExport.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View("Index");
        }

        [HttpPost()]
        public IActionResult Upload(IFormFile file)
        {
            if (file.ContentType == "application/vnd.ms-excel" || file.ContentType == "application/xls")
            {
                using (var stream = file.OpenReadStream())
                {
                    var (occasions, departments) = ImportHelper.ImportDataFromExcelReport(stream);
                    var sourceFilePath = Path.Combine(Environment.CurrentDirectory, "narvarokort-2023-eslovs-kommun.xlsx");
                    var fileExport = ExportHelper.ExportToMunicipalReport(sourceFilePath, occasions, departments);

                    return File(
                        fileContents: fileExport,
                        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileDownloadName: $"narvarokort-{DateTime.Now.Year}-eslovs-kommun.xlsx"
                    );
                }
            }

            return BadRequest();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
