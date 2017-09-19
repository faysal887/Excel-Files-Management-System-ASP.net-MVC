using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using ImportExport.Models;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ImportExport.Controllers
{

    public class ProductController : Controller
    {
        //Product p;
        static List<Product> ExcelData = new List<Product>(); // data posts
        static List<Product> ExcelFileName = new List<Product>();
        static List<Product> _dataFetchedDb = new List<Product>(); // list that stores data fetched from db
        static string fileToOpen = ""; // name of file to open, perform oprations, export
        static List<Product> DeleteMulti = new List<Product>();

        public async Task<ActionResult> Index()
        {
            _dataFetchedDb = await Product.getExcelFileNames();
            return View(_dataFetchedDb);
        }


        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile, FormCollection collection)
        {
            String filename = collection.GetValue("filename").AttemptedValue;

            if (excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select a excel file<br>";
                return View();
            }
            else
            {
                string name = excelfile.FileName;
                if (excelfile != null || excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    String path = Server.MapPath("~/Content/" + filename);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    //Read Data from Excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Product> listProducts = new List<Product>(); /////////////////

                    int count = 0;
                    int c = range.Rows.Count;

                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        Product p = new Product();

                        p.Id = count;
                        p.FirstName = (((Excel.Range)range.Cells[row, 1]).Text);
                        p.LastName = (((Excel.Range)range.Cells[row, 2]).Text);
                        p.Email = (((Excel.Range)range.Cells[row, 3]).Text);
                        p.Title = (((Excel.Range)range.Cells[row, 4]).Text);
                        p.Address1 = (((Excel.Range)range.Cells[row, 5]).Text);
                        p.Address2 = (((Excel.Range)range.Cells[row, 6]).Text);
                        p.City = (((Excel.Range)range.Cells[row, 7]).Text);
                        p.COAddress = (((Excel.Range)range.Cells[row, 8]).Text);
                        p.Zip = (((Excel.Range)range.Cells[row, 9]).Text);
                        p.CompanyEmail = (((Excel.Range)range.Cells[row, 10]).Text);
                        p.CompanyName = (((Excel.Range)range.Cells[row, 11]).Text);
                        p.CompanyLegalName = (((Excel.Range)range.Cells[row, 12]).Text);
                        p.Department = (((Excel.Range)range.Cells[row, 13]).Text);
                        p.CompanyPhone = (((Excel.Range)range.Cells[row, 14]).Text);
                        p.EmployeesMin = (((Excel.Range)range.Cells[row, 15]).Text);
                        p.EmployeesMax = (((Excel.Range)range.Cells[row, 16]).Text);
                        p.CompanyOrganisationNumber = (((Excel.Range)range.Cells[row, 17]).Text);
                        p.VatNumber = (((Excel.Range)range.Cells[row, 18]).Text);
                        p.TurnOver = int.Parse(((Excel.Range)range.Cells[row, 19]).Text);
                        p.County = (((Excel.Range)range.Cells[row, 20]).Text);
                        p.Minicipal = (((Excel.Range)range.Cells[row, 21]).Text);
                        p.SniCode = (((Excel.Range)range.Cells[row, 22]).Text);
                        p.Branch = (((Excel.Range)range.Cells[row, 23]).Text);
                        p.CompanyType = (((Excel.Range)range.Cells[row, 24]).Text);
                        p.Sector = (((Excel.Range)range.Cells[row, 25]).Text);
                        p.MasterSniCode = (((Excel.Range)range.Cells[row, 26]).Text);
                        p.MasterBranch = (((Excel.Range)range.Cells[row, 27]).Text);

                        p.FileName = filename;

                        listProducts.Add(p); //for passing to view
                        ExcelData.Add(p); //to display data before adding to db

                        count++;
                    }


                    //adding column names
                    Product p1 = new Product();
                    p1.FileName = filename;
                    p1.TimeStamp = DateTime.Now;
                    p1.NoOfPosts = range.Rows.Count;
                    ExcelFileName.Add(p1);

                    int n = range.Columns.Count;
                    string[] table = new string[n + 1];

                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        table[i] = ((Excel.Range)range.Cells[1, i]).Text;
                    }

                    ViewBag.ColumnNames = table;
                    ViewBag.ListProducts = listProducts;

                    //selecting name of columns in table product
                    // List<Product> ColumnNames = await Product.GetColumnNamesOfProductTable();
                    //selecting name of columns in table product

                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect<br>";
                    return View("Index");
                }
            }
        }

        public ActionResult addExcelFileToDB()
        {
            Product.addExcelFileDataToDB(ExcelData);
            Product.addExcelFileNameToDB(ExcelFileName);
            return RedirectToAction("Index");
        }

        public async Task<ActionResult> ShowExcelFile(string FileName)
        {
            fileToOpen = FileName;
            _dataFetchedDb = await Product.ShowExcelFile(FileName);
            ViewBag.ListProducts = _dataFetchedDb;
            return View(_dataFetchedDb);
        }

        [HttpPost]
        public async Task<ActionResult> SearchByTurnover(FormCollection collection)
        {

            List<Product> _pb = new List<Product>();
            int val = int.Parse(collection.GetValue("turnover").AttemptedValue);
            string query = collection["filter"].ToString();
            int temp = 0;
            if (query == "greaterthan")
            {
                temp = 1;
                _pb = await Product.searchByTurnover(val, temp);
            }
            else if (query == "lessthan")
            {
                _pb = await Product.searchByTurnover(val, temp);
            }
            ViewBag.ListProducts = _pb;
            return View("ShowExcelFile");
        }

        [HttpPost]
        public ActionResult DeleteMultiple(FormCollection collection)// delete multiple records with checkbox
        {
            string[] ids = collection["productId"].Split(new char[] { ',' });
            Product.DeleteMultipleProducts(ids);
            return RedirectToAction("ShowExcelFile");
        }


        // GET: Product


        public ActionResult DeleteExcels()
        {
            Product.DeleteExcels();
            return RedirectToAction("Index");
        }


        public ActionResult DeleteExcelFile(string FileName)
        {
            Product.DeleteExcelFile(FileName);
            return RedirectToAction("Index");
        }

        [HttpPost]
        public async Task<ActionResult> ExportExcel(FormCollection collection)
        {
            // List<Product> _dataFetchedDb = await Product.ShowExcelFile(fileToOpen);
            string fname = collection.GetValue("filename").AttemptedValue;
            List<Product> datalist = new List<Product>(); // list to store filtered/updated data
            datalist = await Product.ShowExcelFile(fileToOpen);

            try
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                Excel.Range range = worksheet.UsedRange;

                worksheet.Cells[1, 1] = "Id";
                worksheet.Cells[1, 2] = "FirstName ";
                worksheet.Cells[1, 3] = "LastName ";
                worksheet.Cells[1, 4] = "Email ";
                worksheet.Cells[1, 5] = "Title  ";
                worksheet.Cells[1, 6] = "Address1";
                worksheet.Cells[1, 7] = "Address2  ";
                worksheet.Cells[1, 8] = "City  ";
                worksheet.Cells[1, 9] = "COAddress  ";
                worksheet.Cells[1, 10] = "Zip  ";
                worksheet.Cells[1, 11] = "CompanyEmail  ";
                worksheet.Cells[1, 12] = "CompanyName  ";
                worksheet.Cells[1, 13] = "CompanyLegalName  ";
                worksheet.Cells[1, 14] = "Department  ";
                worksheet.Cells[1, 15] = "CompanyPhone  ";
                worksheet.Cells[1, 16] = "EmployeesMin  ";
                worksheet.Cells[1, 17] = "EmployeesMax  ";
                worksheet.Cells[1, 18] = "CompanyOrganisationNumber  ";
                worksheet.Cells[1, 19] = "VatNumber  ";
                worksheet.Cells[1, 20] = "TurnOver  ";
                worksheet.Cells[1, 21] = "County  ";
                worksheet.Cells[1, 22] = "Minicipal  ";
                worksheet.Cells[1, 23] = "SniCode  ";
                worksheet.Cells[1, 24] = "Branch  ";
                worksheet.Cells[1, 25] = "CompanyType  ";
                worksheet.Cells[1, 26] = "Sector  ";
                worksheet.Cells[1, 27] = "MasterSniCode  ";
                worksheet.Cells[1, 28] = "MasterBranch  ";

                int row = 2;
                int count = 0;//
                foreach (var p in datalist)
                {
                    worksheet.Cells[row, 1] = p.Id;
                    worksheet.Cells[row, 2] = p.FirstName;
                    worksheet.Cells[row, 3] = p.LastName;
                    worksheet.Cells[row, 4] = p.Email;
                    worksheet.Cells[row, 5] = p.Title;
                    worksheet.Cells[row, 6] = p.Address1;
                    worksheet.Cells[row, 7] = p.Address2;
                    worksheet.Cells[row, 8] = p.City;
                    worksheet.Cells[row, 9] = p.COAddress;
                    worksheet.Cells[row, 10] = p.Zip;
                    worksheet.Cells[row, 11] = p.CompanyEmail;
                    worksheet.Cells[row, 12] = p.CompanyName;
                    worksheet.Cells[row, 13] = p.CompanyLegalName;
                    worksheet.Cells[row, 14] = p.Department;
                    worksheet.Cells[row, 15] = p.CompanyPhone;
                    worksheet.Cells[row, 16] = p.EmployeesMin;
                    worksheet.Cells[row, 17] = p.EmployeesMax;
                    worksheet.Cells[row, 18] = p.CompanyOrganisationNumber;
                    worksheet.Cells[row, 19] = p.VatNumber;
                    worksheet.Cells[row, 20] = p.TurnOver;
                    worksheet.Cells[row, 21] = p.County;
                    worksheet.Cells[row, 22] = p.Minicipal;
                    worksheet.Cells[row, 23] = p.SniCode;
                    worksheet.Cells[row, 24] = p.Branch;
                    worksheet.Cells[row, 25] = p.CompanyType;
                    worksheet.Cells[row, 26] = p.Sector;
                    worksheet.Cells[row, 27] = p.MasterSniCode;
                    worksheet.Cells[row, 28] = p.MasterBranch;
                    row++;//________________________    
                    count++;
                }


                workbook.SaveAs("f:\\excels\\" + fname + ".xls");

                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                application.Quit();
                Marshal.FinalReleaseComObject(application);

                //updating no of posts after filtering/////////////////////////////

                Product.UpdateNoOfPosts(count);//

            }
            catch (Exception ex)
            {

            }
            return RedirectToAction("Index");
        }
    }
}