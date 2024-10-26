using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
using Irony.Parsing;

namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class OrderController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public OrderController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<OrderModel> orders = new List<OrderModel>
        //{
        //    new OrderModel{OrderID=1,OrderDate=new DateTime(),CustomerID=1,PaymentMode="Online",TotalAmount=10000,ShippingAddress="USA",UserID=1},
        //    new OrderModel{OrderID=2,OrderDate=new DateTime(),CustomerID=2,PaymentMode="Cash",TotalAmount=15000,ShippingAddress="London",UserID=2},
        //    new OrderModel{OrderID=3,OrderDate=new DateTime(),CustomerID=3,PaymentMode="Cheque",TotalAmount=2000,ShippingAddress="US",UserID=3}
        //};
        #endregion
        #region OrderList
        public IActionResult OrderList()
        {
            DataTable table = GetOrderData();
            return View(table);
        }
        #endregion
        #region OrderDelete
        public IActionResult OrderDelete(int OrderID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_OrderDemo_Delete";
                command.Parameters.Add("@OrderID", SqlDbType.Int).Value = OrderID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "Order deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }
            return RedirectToAction("OrderList");
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
        #region CustomerDropDown
        public void CustomerDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection1 = new SqlConnection(connectionString);
            connection1.Open();
            SqlCommand command1 = connection1.CreateCommand();
            command1.CommandType = System.Data.CommandType.StoredProcedure;
            command1.CommandText = "PR_Customer_DropDown";
            SqlDataReader reader1 = command1.ExecuteReader();
            DataTable dataTable1 = new DataTable();
            dataTable1.Load(reader1);
            List<CustomerDropDownModel> customerList = new List<CustomerDropDownModel>();
            foreach (DataRow data in dataTable1.Rows)
            {
                CustomerDropDownModel customerDropDownModel = new CustomerDropDownModel();
                customerDropDownModel.CustomerID = Convert.ToInt32(data["CustomerID"]);
                customerDropDownModel.CustomerName = data["CustomerName"].ToString();
                customerList.Add(customerDropDownModel);
            }
            ViewBag.CustomerList = customerList;
        }
        #endregion
        #region OrderAddEdit
        public IActionResult OrderAddEdit(int OrderID)
        {
            UserDropDown();
            CustomerDropDown();
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDemo_SelectByPK";
            command.Parameters.AddWithValue("@OrderID", OrderID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            OrderModel orderModel = new OrderModel();

            foreach (DataRow dataRow in table.Rows)
            {
                orderModel.OrderID = Convert.ToInt32(@dataRow["OrderID"]);
                orderModel.OrderDate = Convert.ToDateTime(@dataRow["OrderDate"]);
                orderModel.OrderNumber = @dataRow["OrderNumber"].ToString();
                orderModel.CustomerID = Convert.ToInt32(@dataRow["CustomerID"]);
                orderModel.PaymentMode = @dataRow["PaymentMode"].ToString();
                orderModel.TotalAmount = Convert.ToDecimal(dataRow["TotalAmount"]);
                orderModel.ShippingAddress = @dataRow["ShippingAddress"].ToString();
                orderModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }

            return View("OrderAddEdit", orderModel);
        }
        #endregion
        #region OrderSave
        public IActionResult OrderSave(OrderModel orderModel)
        {
            if (orderModel.UserID <= 0)
            {
                ModelState.AddModelError("UserID", "A valid User is required.");
            }

            if (orderModel.CustomerID <= 0)
            {
                ModelState.AddModelError("CustomerID", "A valid Customer is required.");
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
                        if (orderModel.OrderID == null)
                        {
                            command.CommandText = "PR_OrderDemo_Insert";
                        }
                        else
                        {
                            command.CommandText = "PR_OrderDemo_Update";
                            command.Parameters.Add("@OrderID", SqlDbType.Int).Value = orderModel.OrderID;
                        }
                        command.Parameters.Add("@OrderDate", SqlDbType.Date).Value = orderModel.OrderDate;
                        command.Parameters.Add("@OrderNumber", SqlDbType.VarChar).Value = orderModel.OrderNumber;
                        command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = orderModel.CustomerID;
                        command.Parameters.Add("@PaymentMode", SqlDbType.VarChar).Value = orderModel.PaymentMode;
                        command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = orderModel.TotalAmount;
                        command.Parameters.Add("@ShippingAddress", SqlDbType.VarChar).Value = orderModel.ShippingAddress;
                        command.Parameters.Add("@UserID", SqlDbType.Int).Value = orderModel.UserID;

                        command.ExecuteNonQuery();
                        TempData["SuccessMessage"] = orderModel.OrderID == null
                        ? "Order successfully added."
                        : "Order successfully updated.";
                    }
                }

                return RedirectToAction("OrderList");
            }
            UserDropDown();
            CustomerDropDown();
            return View("OrderAddEdit", orderModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetOrderData();

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

                    string fileName = "Orders.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetOrderData
        private DataTable GetOrderData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_OrderDemo_SelectAll";
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
            string orderdate = form["Order.OrderDate"];
            string customerid = form["Order.CustomerID"];
            string paymentMode = form["Order.PaymentMode"];
            string totalAmount = form["Order.TotalAmount"];
            string shippingAddress = form["Order.ShippingAddress"];

            // Store values in ViewBag to display them
            ViewBag.OrderDate = orderdate;
            ViewBag.CustomerID = customerid;
            ViewBag.PaymentMode = paymentMode;
            ViewBag.TotalAmount = totalAmount;
            ViewBag.ShippingAddress = shippingAddress;

            return View();
        }
        #endregion
    }
}
