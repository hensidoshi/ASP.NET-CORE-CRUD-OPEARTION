using CRUD_Demo1.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using System.Data;
using ClosedXML.Excel;

namespace CRUD_Demo1.Controllers
{
    public class UserController : Controller
    {
        #region Configuration
        private IConfiguration configuration;

        public UserController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }
        //public static List<UserModel> users = new List<UserModel>
        //{
        //    new UserModel{UserID= 1,UserName="John", Email="john12@example.com",Password="12345",MobileNo="123-456-7890",Address="USA",IsActive=true},
        //    new UserModel{UserID= 2,UserName="Alice", Email="alice34@example.com",Password="45678",MobileNo="7890-456-123",Address="London",IsActive=true},
        //    new UserModel{UserID= 3,UserName="Bob",Email="bob@example.com",Password="78901",MobileNo="555-666-7777",Address="US",IsActive=true}
        //};
        #endregion
        #region UserList
        public IActionResult UserList()
        {
            DataTable table = GetUserData();
            return View(table);
        }
        #endregion
        #region UserDelete
        public IActionResult UserDelete(int UserID)
        {
            try
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "PR_UserDemo_Delete";
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = UserID;
                command.ExecuteNonQuery();
                TempData["DeleteSuccessMessage"] = "User deleted successfully.";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "A foreign key constraint error occurred. Please try again.";
            }

            return RedirectToAction("UserList");
        }
        #endregion
        #region UserAddEdit
        public IActionResult UserAddEdit(int UserID)
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_UserDemo_SelectByPK";
            command.Parameters.AddWithValue("@UserID", UserID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            UserModel userModel = new UserModel();

            foreach (DataRow dataRow in table.Rows)
            {
                userModel.UserID = Convert.ToInt32(@dataRow["UserID"]);
                userModel.UserName = @dataRow["UserName"].ToString();
                userModel.Email = @dataRow["Email"].ToString();
                userModel.Password = @dataRow["Password"].ToString();
                userModel.MobileNo = @dataRow["MobileNo"].ToString();
                userModel.Address = @dataRow["Address"].ToString();
                userModel.IsActive = Convert.ToBoolean(@dataRow["IsActive"]);
            }

            return View("UserAddEdit", userModel);
        }
        #endregion
        #region UserSave
        public IActionResult UserSave(UserModel userModel)
        {
            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    if (userModel.UserID == null)
                    {
                        command.CommandText = "PR_UserDemo_Insert";
                    }
                    else 
                    {
                        command.CommandText = "PR_UserDemo_Update";
                        command.Parameters.Add("@UserID", SqlDbType.Int).Value = userModel.UserID;
                    }
                    command.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userModel.UserName;
                    command.Parameters.Add("@Email", SqlDbType.VarChar).Value = userModel.Email;
                    command.Parameters.Add("@Password", SqlDbType.VarChar).Value = userModel.Password;
                    command.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = userModel.MobileNo;
                    command.Parameters.Add("@Address", SqlDbType.VarChar).Value = userModel.Address;
                    command.Parameters.Add("@IsActive", SqlDbType.Bit).Value = userModel.IsActive;
                    command.ExecuteNonQuery();
                    TempData["SuccessMessage"] = userModel.UserID == null
                    ? "User successfully added."
                    : "User successfully updated.";
                }
                return RedirectToAction("UserList");
            }

            return View("UserAddEdit", userModel);
        }
        #endregion
        #region ExportToExcel
        public async Task<IActionResult> ExportToExcel()
        {
            DataTable data = GetUserData();

            using (var workbook = new XLWorkbook())
            {
                //Console.WriteLine("workbook:- " + workbook);
                var worksheet = workbook.Worksheets.Add("Users");
                //Console.WriteLine("worksheet:- " + worksheet);
                worksheet.Cell(1, 1).InsertTable(data);
                worksheet.Columns().AdjustToContents();  // Automatically adjust the column widths based on content

                using (var stream = new MemoryStream())
                {
                    //Console.WriteLine("stream:- " + stream);
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    string fileName = "Users.xlsx";
                    Console.WriteLine(stream.ToArray());
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
        }
        #endregion
        #region GetUserData
        private DataTable GetUserData()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_UserDemo_SelectAll";
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
            string username = form["User.UserName"];
            string email = form["User.Email"];
            string password = form["User.Password"];
            string mobileNo = form["User.MobileNo"];
            string address = form["User.Address"];
            string isActive = form["User.IsActive"];

            // Store values in ViewBag to display them
            ViewBag.UserName = username;
            ViewBag.Email = email;
            ViewBag.ProductCode = password;
            ViewBag.MobileNo = mobileNo;
            ViewBag.Address = address;
            ViewBag.IsActive = isActive;

            return View();
        }
        #endregion
        public IActionResult Register()
        {
            return View("Register");
        }
        #region UserRegister
        public IActionResult UserRegister(UserRegisterModel userRegisterModel)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    string connectionString = this.configuration.GetConnectionString("ConnectionString");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.CommandText = "PR_UserDemo_Register";
                    sqlCommand.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userRegisterModel.UserName;
                    sqlCommand.Parameters.Add("@Password", SqlDbType.VarChar).Value = userRegisterModel.Password;
                    sqlCommand.Parameters.Add("@Email", SqlDbType.VarChar).Value = userRegisterModel.Email;
                    sqlCommand.Parameters.Add("@MobileNo", SqlDbType.VarChar).Value = userRegisterModel.MobileNo;
                    sqlCommand.Parameters.Add("@Address", SqlDbType.VarChar).Value = userRegisterModel.Address;
                    sqlCommand.ExecuteNonQuery();
                    return RedirectToAction("Login", "User");
                }
            }
            catch (Exception e)
            {
                TempData["ErrorMessage"] = e.Message;
                return RedirectToAction("Register");
            }
            return RedirectToAction("Register");
        }
        #endregion
        public IActionResult Login()
        {
            return View("Login");
        }
        #region UserLogin
        public IActionResult UserLogin(UserLoginModel userLoginModel)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    string connectionString = this.configuration.GetConnectionString("ConnectionString");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.CommandText = "PR_UserDemo_Login";
                    sqlCommand.Parameters.Add("@UserName", SqlDbType.VarChar).Value = userLoginModel.UserName;
                    sqlCommand.Parameters.Add("@Password", SqlDbType.VarChar).Value = userLoginModel.Password;
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    DataTable dataTable = new DataTable();
                    dataTable.Load(sqlDataReader);
                    if (dataTable.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dataTable.Rows)
                        {
                            HttpContext.Session.SetString("UserID", dr["UserID"].ToString());
                            HttpContext.Session.SetString("UserName", dr["UserName"].ToString());
                        }

                        return RedirectToAction("Index", "Home");
                    }
                    else
                    {
                        return RedirectToAction("Login", "User");
                    }

                }
            }
            catch (Exception e)
            {
                TempData["ErrorMessage"] = e.Message;
            }

            return RedirectToAction("Login");
        }
        #endregion
       
        #region Logout
        public IActionResult Logout()
        {
            HttpContext.Session.Clear();
            return RedirectToAction("Login", "User");
        }
        #endregion
    }
}
