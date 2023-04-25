using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using TopTenErrors.Services;

namespace TopTenErrors.Controllers;

public class HomeController : Controller
{
    TxtFileServices _txtFileServices;
    public HomeController(TxtFileServices txtFileServices)
    {
        _txtFileServices = txtFileServices;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> ReadTxtFiles(IFormFileCollection TopTenTxtFiles)
    {

        if (_txtFileServices.IsFilesExists(TopTenTxtFiles) == false)
        {
            return BadRequest("No Files were uploaded");
        }

        if (_txtFileServices.IsValidFilesExtentions(TopTenTxtFiles) == false)
        {
            return BadRequest("One or more files not a txt file extentions");
        }
        var listFiles = _txtFileServices.TxtFilter(TopTenTxtFiles);
        await _txtFileServices.CreateExcelSheet(listFiles);

        return RedirectToAction("Index");
    }

}

