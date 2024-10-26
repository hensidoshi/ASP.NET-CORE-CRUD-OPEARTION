using Microsoft.AspNetCore.Mvc;
using CRUD_Demo1.Models;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class ProductController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public ProductController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<ProductModel> products = new List<ProductModel>
        //{
        //    new ProductModel{ProductID = 1,ProductName="Laptop", ProductPrice=75000,ProductCode="101",Description="Electrnoic device",UserID=1},
        //    new ProductModel{ProductID = 2,ProductName="Mobile", ProductPrice=40000,ProductCode="102",Description="Smart Phone",UserID=1},
        //    new ProductModel{ProductID = 3,ProductName="Airbuds", ProductPrice=2000,ProductCode="103",Description="Hearing device",UserID=1}
        //};
        #endregion
        #region ProductList
        public IActionResult ProductList()
        {
            DataTable table = GetProductData();
            return View(table);
        }
        #endregion 
        #region ProductDelete
        public IActionResult ProductDelete(int ProductID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_Product_Delete";
                command.Parameters.Add("@ProductID", SqlDbType.Int).Value = ProductID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "Product deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }

            return RedirectToAction("ProductList");
        }
        #endregion
        #region UserDropDown
        public void UserDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection1 = new SqlConnection(connectionString);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = System.Data.CommandType.StoredProcedure;
            command1.CommandText = "PR_UserDemo_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            connection1.Close();

            List<UserDropDownModel> users = new List<UserDropDownModel>();
            foreach (DataRow dataRow in dataTable1.Rows)
            {
                UserDropDownModel userDropDownModel = new UserDropDownModel();
                userDropDownModel.UserID = Convert.ToInt32(dataRow["UserID"]);
                userDropDownModel.UserName = dataRow["UserName"].ToString();
                users.Add(userDropDownModel);
            }
            ViewBag.UserList = users;
        }
        #endregion
        public IActionResult ProductDetail(int ProductID)
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Product_SelectByPK";
            command.Parameters.AddWithValue("@ProductID", ProductID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            ProductModel productModel = new ProductModel();

            foreach (DataRow dr in table.Rows)
            {
                productModel.ProductID = Convert.ToInt32(@dr["ProductID"]);
                productModel.ProductName = @dr["ProductName"].ToString();
                productModel.ProductCode = @dr["ProductCode"].ToString();
                productModel.ProductPrice = Convert.ToDecimal(@dr["ProductPrice"]);
                productModel.Description = @dr["Description"].ToString();
                productModel.UserID = Convert.ToInt32(@dr["UserID"]);
            }
            return View(ProductDetail);
        }
        #region ProductAddEdit
        public IActionResult ProductAddEdit(int ProductID)
        {
            UserDropDown();

            #region ProductByID
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Product_SelectByPK";
            command.Parameters.AddWithValue("@ProductID", ProductID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            ProductModel productModel = new ProductModel();

            foreach (DataRow dr in table.Rows)
            {
                productModel.ProductID = Convert.ToInt32(@dr["ProductID"]);
                productModel.ProductName = @dr["ProductName"].ToString();
                productModel.ProductCode = @dr["ProductCode"].ToString();
                productModel.ProductPrice = Convert.ToDecimal(@dr["ProductPrice"]);
                productModel.Description = @dr["Description"].ToString();
                productModel.UserID = Convert.ToInt32(@dr["UserID"]);
            }

            #endregion

            return View("ProductAddEdit", productModel);
        }
        #endregion
        #region ProductSave
        public IActionResult ProductSave(ProductModel productModel)
        {
            if (productModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }

            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                if (productModel.ProductID == null)
                {
                    command.CommandText = "PR_Product_Insert";
                }
                else
                {
                    command.CommandText = "PR_Product_Update";
                    command.Parameters.Add("@ProductID", SqlDbType.Int).Value = productModel.ProductID;
                }
                command.Parameters.Add("@ProductName", SqlDbType.VarChar).Value = productModel.ProductName;
                command.Parameters.Add("@ProductCode", SqlDbType.VarChar).Value = productModel.ProductCode;
                command.Parameters.Add("@ProductPrice", SqlDbType.Decimal).Value = productModel.ProductPrice;
                command.Parameters.Add("@Description", SqlDbType.VarChar).Value = productModel.Description;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = productModel.UserID;
                command.ExecuteNonQuery();

                TempData["SuccessMessage"] = productModel.ProductID == null
                ? "Product successfully added."
                : "Product successfully updated.";

                return RedirectToAction("ProductList");
            }
            UserDropDown();
            return View("ProductAddEdit", productModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetProductData();

            using (var workbook = new XLWorkbook())
            {
                //Console.WriteLine("workbook:- " + workbook);
                var worksheet = workbook.Worksheets.Add("Products");
                //Console.WriteLine("worksheet:- " + worksheet);
                worksheet.Cell(1, 1).InsertTable(data);
                worksheet.Columns().AdjustToContents();  // Automatically adjust the column widths based on content

                using (var stream = new MemoryStream())
                {
                    //Console.WriteLine("stream:- " + stream);
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    string fileName = "Products.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetProductData
        private DataTable GetProductData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Product_SelectAll";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return table;
        }
        #endregion
        [HttpPost]
        #region Submit
        public ActionResult Submit(IFormCollection form)
        {
            // Extract values from the form
            string productName = form["Product.ProductName"];
            string productPrice = form["Product.ProductPrice"];
            string productCode = form["Product.ProductCode"];
            string description = form["Product.Description"];

            // Store values in ViewBag to display them
            ViewBag.ProductName = productName;
            ViewBag.ProductPrice = productPrice;
            ViewBag.ProductCode = productCode;
            ViewBag.Description = description;

            return View();
        }
        #endregion

    }
}
