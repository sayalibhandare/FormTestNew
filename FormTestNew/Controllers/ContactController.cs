using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Xml.Linq;
using FormTestNew.Models;
using IronPdf;
using IronXL;


namespace FormTestNew.Controllers
{
    public class ContactController : Controller
    {
        // GET: Contact/Create
        public ActionResult Index()
        {
          

            using (SqlConnection connection = new SqlConnection("Server=NODE191\\SQLEXPRESS;Initial Catalog=FormTest;Integrated Security=True"))
            {

                String selectQuery = "SELECT * FROM dbo.FormTable";

                // create object for Datatable
                var table = new DataTable();

                // Acquire all data from databse and fill it into Datatable
                using (var dataTable = new SqlDataAdapter(selectQuery, connection))
                {
                    dataTable.Fill(table);
                }

                // Create an empty excel sheet using IronXL
                WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);

                // Add authore
                wb.Metadata.Author = "Sayali";

                WorkSheet ws = wb.DefaultWorkSheet;
                int rowCount = 1;

                // Write all data from databse on excel sheet
                foreach (DataRow row in table.Rows)
                {
                    ws["A" + (rowCount)].Value = row[4].ToString();
                    ws["B" + (rowCount)].Value = row[0].ToString();
                    ws["C" + (rowCount)].Value = row[1].ToString();
                    ws["D" + (rowCount)].Value = row[2].ToString();
                    ws["E" + (rowCount)].Value = row[3].ToString();
                    rowCount++;
                }

                // Save file as CSV on below address
                
                wb.SaveAsCsv("C:/Users/sayali.sarode/sayali-sample.csv");


            }
            ViewBag.Message = "The excel file has been saved on your PC of location C:/Users/sayali.sarode/sayali-sample.CSV";

            return View();
        }

        // GET: Contact/Create
        public ActionResult Create()
        {

           return View();
        }

        // POST: Contact/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
               
                string provider = "System.Data.SqlClient"; 
                DbProviderFactory factory = DbProviderFactories.GetFactory(provider);
                List<FormDataModel> myUsers = new List<FormDataModel>();

                // Connect to databse
                using (SqlConnection connection = new SqlConnection("Server=NODE191\\SQLEXPRESS;Initial Catalog=FormTest;Integrated Security=True"))
                {
                    String query = "INSERT INTO dbo.FormTable (Name,Dob,Phone,Email) VALUES (@name,@dob,@phone,@email)";

                    // Create a SQL query
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Add all properies from view onto FormData table
                        command.Parameters.AddWithValue("@name", collection["Name"]);
                        command.Parameters.AddWithValue("@dob", collection["DOB"]);
                        command.Parameters.AddWithValue("@phone", collection["Phone"]);
                        command.Parameters.AddWithValue("@email", collection["Email"]);

                        // Open database connection
                        connection.Open();
                        //Execute SQL query
                        int result = command.ExecuteNonQuery();

                        // Check Error
                        if (result < 0)
                            Console.WriteLine("Error inserting data into Database!");
                    }

                    // Call get All function which return all rows from database
                    getAll();

                    // Return all valuses to view
                    return View("Index");
                }
            }
            catch
            {
                return View();
            }
        }
        public ActionResult getAll()
        {
            
            List<FormDataModel> viewUsers = getUsers();
            ViewBag.users = viewUsers;          
            return View("Index");
        }
       
        public List<FormDataModel> getUsers()
        {
            List<FormDataModel> myUsers = new List<FormDataModel>();

            // Connect to database
            using (SqlConnection connection = new SqlConnection("Server=NODE191\\SQLEXPRESS;Initial Catalog=FormTest;Integrated Security=True"))
            {
                
                String selectQuery = "SELECT * FROM dbo.FormTable";
                var table = new DataTable();

                // Create sql query object
                using (SqlCommand command = new SqlCommand(selectQuery, connection))
                {
                   // open db connection
                    connection.Open();

                    // Execute query to select all data from database
                    int result = command.ExecuteNonQuery();

                    // create object of SqlDataReader
                    SqlDataReader rdr = command.ExecuteReader();

                    // Read all data and add to each user property using FormDataModel object
                    while (rdr.Read())
                    {
                        FormDataModel retrievedUser = new FormDataModel();

                        retrievedUser.Name = rdr[0].ToString();
                        retrievedUser.DOB = rdr[1].ToString();
                        retrievedUser.Email = rdr[2].ToString();
                        retrievedUser.Phone = rdr[3].ToString();
                        retrievedUser.Id = rdr[4].ToString();

                        myUsers.Add(retrievedUser);
                    }
                }
                return myUsers;
            }
        }

    }
}
