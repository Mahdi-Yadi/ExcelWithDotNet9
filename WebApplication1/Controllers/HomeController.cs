using System.Text;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
namespace WebApplication1.Controllers;
public class HomeController : Controller
{

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

}