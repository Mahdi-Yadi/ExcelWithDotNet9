using System.Text;
using ExcelDataReader;
using ExcelWithDotNet9.Web.Data.Contexts;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace WebApplication1.Controllers;
public class HomeController : Controller
{

    private readonly DBContext _dbContext;

    public HomeController(DBContext dbContext)
    {
        _dbContext = dbContext;
    }

    public IActionResult Index()
    {

        return View();
    }


    [HttpPost]
    public IActionResult ImportExecl(IFormFile execlFile)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (execlFile != null)
        {

            var uploadFolder =
                $"{Directory.GetCurrentDirectory()}\\wwwroot\\files\\";

            var filePath = Path.Combine(uploadFolder,
                $"{DateTime.Now.Hour.ToString().Replace(" ", "-")}" + execlFile.FileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                execlFile.CopyTo(stream);
            }

            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        bool isheaderskipped =  false;

                        while (reader.Read())
                        {
                            if (!isheaderskipped)
                            {
                                isheaderskipped= true;
                                continue;
                            }

                            // save to database

                            ViewBag.data1 = reader.GetValue(0).ToString();
                            ViewBag.data2 = reader.GetValue(1).ToString();
                            return Redirect("/");
                        }

                    } while (reader.NextResult());
                }

            }

        }

        return Redirect("/");
    }


    [HttpGet]
    public IActionResult ExportExecl()
    {
        var db = _dbContext.Order.ToList();

        if (db.Count != 0)
        {
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var work = package.Workbook.Worksheets.Add("Orders");

                work.Cells[1, 1].Value = "Id";
                work.Cells[1, 2].Value = "جمع فاکتور";

                int row = 2;

                foreach (var item in db)
                {
                    work.Cells[row, 1].Value = item.Id.ToString();
                    work.Cells[row, 2].Value = item.Sum;
                    row++;
                }
                package.Save();
            }

            stream.Position = 0;
            string excelName = $"Orders-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

        return Redirect("/");
    }


}