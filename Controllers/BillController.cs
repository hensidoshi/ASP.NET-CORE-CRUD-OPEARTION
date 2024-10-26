using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
using Irony.Parsing;

namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class BillController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public BillController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<BillModel> bills = new List<BillModel>
        //{
        //    new BillModel{BillID=1,BillNumber="101",BillDate=new DateTime(),TotalAmount=10000,Discount=1000,NetAmount=9000,UserID=1},
        //    new BillModel{BillID=2,BillNumber="102",BillDate=new DateTime(),TotalAmount=15000,Discount=1500,NetAmount=13500,UserID=2},
        //    new BillModel{BillID=3,BillNumber="103",BillDate=new DateTime(),TotalAmount=2000,Discount=100,NetAmount=1900,UserID=3}
        //};
        #endregion
        #region BillList
        public IActionResult BillList()
        {
           DataTable table = GetBillData();
            return View(table);
        }
        #endregion
        #region BillDelete
        public IActionResult BillDelete(int BillID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_Bills_Delete";
                command.Parameters.Add("@BillID", SqlDbType.Int).Value = BillID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "Bill deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }

            return RedirectToAction("BillList");
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
        #region BillAddEdit
        public IActionResult BillAddEdit(int BillID)
        {
            OrderDropDown();
            UserDropDown();

            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Bills_SelectByPK";
            command.Parameters.AddWithValue("@BillID", BillID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            BillModel billModel = new BillModel();

            foreach (DataRow dataRow in table.Rows)
            {
                billModel.BillID = Convert.ToInt32(@dataRow["BillID"]);
                billModel.BillNumber = @dataRow["BillNumber"].ToString();
                billModel.BillDate = Convert.ToDateTime(dataRow["BillDate"]);
                billModel.OrderID = Convert.ToInt32(@dataRow["OrderID"]);
                billModel.TotalAmount = Convert.ToDecimal(dataRow["TotalAmount"]);
                billModel.Discount = Convert.ToDecimal(dataRow["Discount"]);
                billModel.NetAmount = Convert.ToDecimal(dataRow["NetAmount"]);
                billModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }

            return View("BillAddEdit", billModel);
        }
        #endregion
        #region BillSave
        public IActionResult BillSave(BillModel billModel)
        {
            if (billModel.UserID <= 0)
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
                if (billModel.BillID == null)
                {
                    command.CommandText = "PR_Bills_Insert";
                }
                else
                {
                    command.CommandText = "PR_Bills_Update";
                    command.Parameters.Add("@BillID", SqlDbType.Int).Value = billModel.BillID;
                }
                command.Parameters.Add("@BillNumber", SqlDbType.VarChar).Value = billModel.BillNumber;
                command.Parameters.Add("@BillDate", SqlDbType.Date).Value = billModel.BillDate;
                command.Parameters.Add("@OrderID", SqlDbType.Int).Value = billModel.OrderID;
                command.Parameters.Add("@TotalAmount", SqlDbType.Decimal).Value = billModel.TotalAmount;
                command.Parameters.Add("@Discount", SqlDbType.Decimal).Value = billModel.Discount;
                command.Parameters.Add("@NetAmount", SqlDbType.Decimal).Value = billModel.NetAmount;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = billModel.UserID;
                command.ExecuteNonQuery();
                TempData["SuccessMessage"] = billModel.BillID == null
                ? "Bill successfully added."
                : "Bill successfully updated.";
                return RedirectToAction("BillList");
            }
            OrderDropDown();
            UserDropDown();
            return View("BillAddEdit", billModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetBillData();

            using (var workbook = new XLWorkbook())
            {
                //Console.WriteLine("workbook:- " + workbook);
                var worksheet = workbook.Worksheets.Add("Bill");
                //Console.WriteLine("worksheet:- " + worksheet);
                worksheet.Cell(1, 1).InsertTable(data);
                worksheet.Columns().AdjustToContents();  // Automatically adjust the column widths based on content

                using (var stream = new MemoryStream())
                {
                    //Console.WriteLine("stream:- " + stream);
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    string fileName = "Bills.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetBillData
        private DataTable GetBillData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Bills_SelectAll";
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
            string billNumber = form["Bill.BillNumber"];
            string billDate = form["Bill.BillDate"];
            string orderid = form["Bill.OrderID"];
            string totalAmount = form["Bill.TotalAmount"];
            string discount = form["Bill.Discount"];
            string netAmount = form["Bill.NetAmount"];

            // Store values in ViewBag to display them
            ViewBag.BillNumber = billNumber;
            ViewBag.BillDate = billDate;
            ViewBag.OrderID = orderid;
            ViewBag.TotalAmount = totalAmount;
            ViewBag.Discount = discount;
            ViewBag.NetAmount = netAmount;

            return View();
        }
        #endregion
    }
}
