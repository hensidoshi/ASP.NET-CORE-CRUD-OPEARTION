using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;

namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class OrderDetailController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public OrderDetailController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<OrderDetailModel> orderdetails = new List<OrderDetailModel>
        //{
        //    new OrderDetailModel{OrderDetailID=1,OrderID=1,ProductID=1,Quantity=10,Amount=1000,TotalAmount=10000,UserID=1},
        //    new OrderDetailModel{OrderDetailID=2,OrderID=2,ProductID=2,Quantity=5,Amount=1000,TotalAmount=15000,UserID=2},
        //    new OrderDetailModel{OrderDetailID=3,OrderID=3,ProductID=3,Quantity=7,Amount=1200,TotalAmount=2000,UserID=3}
        //};
        #endregion
        #region OrderDetailsList
        public IActionResult OrderDetailsList()
        {
            DataTable table = GetOrderDetailData();
            return View(table);
        }
        #endregion
        #region OrderDetailDelete
        public IActionResult OrderDetailDelete(int OrderDetailID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_OrderDetail_Delete";
                command.Parameters.Add("@OrderDetailID", SqlDbType.Int).Value = OrderDetailID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "OrderDetail deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }
            return RedirectToAction("OrderDetailsList");
        }
        #endregion
        #region OrderDropDown
        public void OrderDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection1 = new SqlConnection(connectionString);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = System.Data.CommandType.StoredProcedure;
            command1.CommandText = "PR_OrderDemo_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            connection1.Close();

            List<OrderDropDownModel> orders = new List<OrderDropDownModel>();
            foreach (DataRow dataRow in dataTable1.Rows)
            {
                OrderDropDownModel orderDropDownModel = new OrderDropDownModel();
                orderDropDownModel.OrderID = Convert.ToInt32(dataRow["OrderID"]);
                orderDropDownModel.OrderNumber = dataRow["OrderNumber"].ToString();
                orders.Add(orderDropDownModel);
            }
            ViewBag.OrderList = orders;
        }
        #endregion
        #region ProductDropDown
        public void ProductDropDown()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();

            SqlCommand commandProduct = connection.CreateCommand();
            commandProduct.CommandType = CommandType.StoredProcedure;
            commandProduct.CommandText = "PR_Product_DropDown";
            SqlDataReader readerCustomer = commandProduct.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(readerCustomer);
            readerCustomer.Close();

            List<ProductDropDownModel> ProductList = new List<ProductDropDownModel>();
            foreach (DataRow dr in table.Rows)
            {
                ProductDropDownModel model = new ProductDropDownModel();
                model.ProductID = Convert.ToInt32(dr["ProductID"]);
                model.ProductName = dr["ProductName"].ToString();
                ProductList.Add(model);
            }
            ViewBag.ProductList = ProductList;
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
        #region OrderDetailAddEdit
        public IActionResult OrderDetailAddEdit(int OrderDetailID)
        {
            OrderDropDown();
            ProductDropDown();
            UserDropDown();

            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_SelectByPK";
            command.Parameters.AddWithValue("@OrderDetailID", OrderDetailID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            OrderDetailModel orderDetailModel = new OrderDetailModel();

            foreach (DataRow dataRow in table.Rows)
            {
                orderDetailModel.OrderDetailID = Convert.ToInt32(@dataRow["OrderDetailID"]);
                orderDetailModel.OrderID = Convert.ToInt32(@dataRow["OrderID"]);
                orderDetailModel.ProductID = Convert.ToInt32(@dataRow["ProductID"]);
                orderDetailModel.Quantity = Convert.ToInt32(@dataRow["Quantity"]);
                orderDetailModel.Amount = Convert.ToDecimal(dataRow["Amount"]);
                orderDetailModel.TotalAmount = Convert.ToDecimal(dataRow["TotalAmount"]);
                orderDetailModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }

            return View("OrderDetailAddEdit", orderDetailModel);
        }
        #endregion
        #region OrderDetailSave
        public IActionResult OrderDetailSave(OrderDetailModel orderDetailModel)
        {
            if (orderDetailModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }
            if (orderDetailModel.ProductID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid Product is required.");
            }
            if (orderDetailModel.OrderID <= 0)
            {
                ModelState.AddModelError("OrderID", "A valid OrderID is required.");
            }
            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        if (orderDetailModel.OrderDetailID == null)
                        {
                            command.CommandText = "PR_OrderDetail_Insert";
                        }
                        else
                        {
                            command.CommandText = "PR_OrderDetail_Update";
                            command.Parameters.Add("@OrderDetailID", SqlDbType.Int).Value = orderDetailModel.OrderDetailID;
                        }

                        command.Parameters.Add("@OrderID", SqlDbType.Int).Value = orderDetailModel.OrderID;
                        command.Parameters.Add("@ProductID", SqlDbType.Int).Value = orderDetailModel.ProductID;
                        command.Parameters.Add("@Quantity", SqlDbType.Int).Value = orderDetailModel.Quantity;
                        command.Parameters.Add("@Amount", SqlDbType.Decimal).Value = orderDetailModel.Amount;
                        command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = orderDetailModel.TotalAmount;
                        command.Parameters.Add("@UserID", SqlDbType.Int).Value = orderDetailModel.UserID;

                        command.ExecuteNonQuery();
                        TempData["SuccessMessage"] = orderDetailModel.OrderDetailID == null
                        ? "OrderDetail successfully added."
                        : "OrderDetail successfully updated.";
                    }
                }
                OrderDropDown();
                ProductDropDown();
                UserDropDown();
                return RedirectToAction("OrderDetailsList");
            }

            return View("OrderDetailAddEdit", orderDetailModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetOrderDetailData();

            using (var workbook = new XLWorkbook())
            {
                //Console.WriteLine("workbook:- " + workbook);
                var worksheet = workbook.Worksheets.Add("Orders");
                //Console.WriteLine("worksheet:- " + worksheet);
                worksheet.Cell(1, 1).InsertTable(data);
                worksheet.Columns().AdjustToContents();  // Automatically adjust the column widths based on content

                using (var stream = new MemoryStream())
                {
                    //Console.WriteLine("stream:- " + stream);
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    string fileName = "OrderDetails.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetOrderDetailData
        private DataTable GetOrderDetailData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDetail_SelectAll";
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
            string orderid = form["OrderDetail.OrderID"];
            string productid = form["OrderDetail.ProductID"];
            string quantity = form["OrderDetail.Quantity"];
            string amount = form["OrderDetail.Amount"];
            string totalAmount = form["OrderDetail.TotalAmount"];

            // Store values in ViewBag to display them
            ViewBag.OrderDate = orderid;
            ViewBag.ProductID = productid;
            ViewBag.Quantity = quantity;
            ViewBag.Amount = amount;
            ViewBag.TotalAmount = totalAmount;

            return View();
        }
        #endregion
    }
}

