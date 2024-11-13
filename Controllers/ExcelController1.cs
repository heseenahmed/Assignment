
using Assignment.Models;
using Microsoft.AspNetCore.Mvc;

public class ExcelController : Controller
{
    private readonly ExcelService _excelService;

    public ExcelController()
    {
        _excelService = new ExcelService();
    }

    [HttpGet]
    public IActionResult Index(decimal? total = null)
    {
        return View();
    }

   
    [HttpPost]
    public IActionResult Import(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            ModelState.AddModelError("File", "Please upload a valid Excel file.");
            return View("Index");
        }

        List<TaxRecord> records;
        using (var stream = file.OpenReadStream())
        {
            records = _excelService.ImportExcelFile(stream);
        }

        decimal totalAfterTaxing = records.Sum(r => r.TotalValueAfterTaxing);

        var modifiedSheet = _excelService.AddColumnAndCalculateTotals(records);

        TempData["ModifiedExcel"] = Convert.ToBase64String(modifiedSheet);

        ViewBag.TotalAfterTaxing = totalAfterTaxing;

        return View("Index");
    }
    [HttpGet]
    public IActionResult DownloadModifiedExcel()
    {
        if (TempData["ModifiedExcel"] != null)
        {
            byte[] fileBytes = Convert.FromBase64String(TempData["ModifiedExcel"].ToString());
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ModifiedSheet.xlsx");
        }
        return RedirectToAction("Index");
    }


}
