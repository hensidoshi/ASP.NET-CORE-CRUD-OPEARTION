using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;
using Irony.Parsing;

namespace CRUD_Demo1.Controllers
{
    [CheckAccess]
    public class CustomerController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public CustomerController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<CustomerModel> customers = new List<CustomerModel>
        //{
        //    new CustomerModel{CustomerID=1,CustomerName="Job",HomeAddress="USA",Email="job12@example.com",MobileNo="123-456-7890",GSTNO="123456",CityName="New York",Pincode="121212",NetAmount=9000,UserID=1},
        //    new CustomerModel{CustomerID=2,CustomerName="Alice",HomeAddress="London",Email="alice34@example.com",MobileNo="7890-456-123",GSTNO="456789",CityName="Paris",Pincode="343434",NetAmount=13500,UserID=2},
        //    new CustomerModel{CustomerID=3,CustomerName="Bob",HomeAddress="US",Email="bob@example.com",MobileNo="555-666-7777",GSTNO="789012",CityName="Houston",Pincode="565656",NetAmount=1900,UserID=3}
        //};
        #endregion
        #region CustomerList
        public IActionResult CustomerList()
        {
            DataTable table = GetCustomerData();
            return View(table);
        }
        #endregion
        #region CustomerDelete
        public IActionResult CustomerDelete(int CustomerID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_Customer_Delete";
                command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = CustomerID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "Customer deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }

            return RedirectToAction("CustomerList");
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
        #region CustomerAddEdit
        public IActionResult CustomerAddEdit(int CustomerID)
        {
            UserDropDown();
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Customer_SelectByPK";
            command.Parameters.AddWithValue("@CustomerID", CustomerID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            CustomerModel customerModel = new CustomerModel();

            foreach (DataRow dataRow in table.Rows)
            {
                customerModel.CustomerID = Convert.ToInt32(@dataRow["CustomerID"]);
                customerModel.CustomerName = @dataRow["CustomerName"].ToString();
                customerModel.HomeAddress = @dataRow["HomeAddress"].ToString();
                customerModel.Email = dataRow["Email"].ToString();
                customerModel.MobileNo = @dataRow["MobileNo"].ToString();
                customerModel.GST_NO = @dataRow["GST_NO"].ToString();
                customerModel.CityName = @dataRow["CityName"].ToString();
                customerModel.Pincode = @dataRow["PinCode"].ToString();
                customerModel.NetAmount = Convert.ToDecimal(dataRow["NetAmount"]);
                customerModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
            }
            return View("CustomerAddEdit", customerModel);
        }
        #endregion
        #region CustomerSave
        public IActionResult CustomerSave(CustomerModel customerModel)
        {
            if (customerModel.UserID <= 0)
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
                if (customerModel.CustomerID == null)
                {
                    command.CommandText = "PR_Customer_Insert";
                }
                else
                {
                    command.CommandText = "PR_Customer_Update";
                    command.Parameters.Add("@CustomerID", SqlDbType.Int).Value = customerModel.CustomerID;
                }
                command.Parameters.Add("@CustomerName", SqlDbType.VarChar).Value = customerModel.CustomerName;
                command.Parameters.Add("@HomeAddress", SqlDbType.VarChar).Value = customerModel.HomeAddress;
                command.Parameters.Add("@Email", SqlDbType.VarChar).Value = customerModel.Email;
                command.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = customerModel.MobileNo;
                command.Parameters.Add("@GST_NO", SqlDbType.VarChar).Value = customerModel.GST_NO;
                command.Parameters.Add("@CityName", SqlDbType.VarChar).Value = customerModel.CityName;
                command.Parameters.Add("@PinCode", SqlDbType.VarChar).Value = customerModel.Pincode;
                command.Parameters.Add("@NetAmount", SqlDbType.Decimal).Value = customerModel.NetAmount;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = customerModel.UserID;
                command.ExecuteNonQuery();
                TempData["SuccessMessage"] = customerModel.CustomerID == null
                    ? "Customer successfully added."
                    : "Customer successfully updated.";
                return RedirectToAction("CustomerList");
            }
            UserDropDown();
            return View("CustomerAddEdit", customerModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetCustomerData();

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

                    string fileName = "Customers.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetCustomerData
        private DataTable GetCustomerData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Customer_SelectAll";
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
            string customerName = form["Customer.CustomerName"];
            string homeAddress = form["Bill.HomeAddress"];
            string email = form["Bill.Email"];
            string mobileNo = form["Bill.MobileNo"];
            string GSTNO = form["Bill.GSTNO"];
            string cityName = form["Bill.CityName"];
            string pinCode = form["Bill.PinCode"];
            string netAmount = form["Bill.NetAmount"];

            // Store values in ViewBag to display them
            ViewBag.CustomerName = customerName;
            ViewBag.HomeAddress = homeAddress;
            ViewBag.Email = email;
            ViewBag.MobileNo = mobileNo;
            ViewBag.GSTNO = GSTNO;
            ViewBag.CityName = cityName;
            ViewBag.PinCode = pinCode;
            ViewBag.NetAmount = netAmount;

            return View();
        }
        #endregion
    }
}
