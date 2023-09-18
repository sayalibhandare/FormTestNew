using IronXL;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FormTestNew.Controllers
{
    public class CSVController : Controller
    {
        // GET: CSV
        public ActionResult Index()
        {
            return View();
        }

        // GET: CSV/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: CSV/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: CSV/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection("Server=NODE191\\SQLEXPRESS;Initial Catalog=FormTest;Integrated Security=True"))
                {

                    String selectQuery = "SELECT * FROM dbo.FormTable";
                    var table = new DataTable();

                    using (var dataTable = new SqlDataAdapter(selectQuery, connection))
                    {
                        dataTable.Fill(table);
                    }
                    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
                    wb.Metadata.Author = "Sayali";
                    WorkSheet ws = wb.DefaultWorkSheet;
                    int rowCount = 1;
                    foreach (DataRow row in table.Rows)
                    {
                        ws["A" + (rowCount)].Value = row[0].ToString();
                        rowCount++;
                    }
                    wb.SaveAsCsv("Save_DataTable_CSV.csv", ";"); // Saved as : Save_DataTable_CSV.Sheet1.csv


                }

                return View("CSV/CSV");
            }
            catch
            {
                return View();
            }
        }

        // GET: CSV/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: CSV/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: CSV/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: CSV/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
