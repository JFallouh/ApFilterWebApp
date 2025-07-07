// Created by James Fallouh
// Date: 2025-07-07

using Microsoft.AspNetCore.Mvc;
using ApFilterWebApp.Services;

namespace ApFilterWebApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelFilterService _svc;
        public HomeController(ExcelFilterService svc) => _svc = svc;

        [HttpGet]
        public IActionResult Index() => View();

        [HttpPost]
        public IActionResult Index(IFormFile? sourceFile, string? destinationFolder)
        {
            try
            {
                _svc.Process(sourceFile?.OpenReadStream(), destinationFolder);
                ViewBag.Message = "Files generated successfully.";
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
            }
            return View();
        }
    }
}
