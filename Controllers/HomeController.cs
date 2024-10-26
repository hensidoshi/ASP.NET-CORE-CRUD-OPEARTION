using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Employee()
        {
            return View();
        }
        public IActionResult EmployeeForm()
        {
            return View();
        }
        public IActionResult Department()
        {
            return View();
        }
        public IActionResult DepartmentForm()
        {
            return View();
        }
        public IActionResult Project()
        {
            return View();
        }
        public IActionResult ProjectForm()
        {
            return View();
        }
        public IActionResult EmployeeProject()
        {
            return View();
        }
        public IActionResult EmployeeProjectForm()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
