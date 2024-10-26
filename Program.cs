var builder = WebApplication.CreateBuilder(args);

// Enable session
builder.Services.AddDistributedMemoryCache();
builder.Services.AddHttpContextAccessor(); // Allows HttpContext access
builder.Services.AddSession(); // Adds session services

builder.Services.AddControllersWithViews();

var app = builder.Build();

// Use session middleware
app.UseSession();
app.UseStaticFiles();

// Use routing and other middleware
app.UseRouting();
    app.MapControllerRoute(
        name: "default",
        pattern: "{controller=Home}/{action=Index}/{id?}");


app.Run();


//var builder = WebApplication.CreateBuilder(args);

//// Add services to the container
//builder.Services.AddDistributedMemoryCache();
//builder.Services.AddSession(options =>
//{
//    options.IdleTimeout = TimeSpan.FromMinutes(30);
//    options.Cookie.HttpOnly = true;
//    options.Cookie.IsEssential = true;
//});

//// Add MVC support
//builder.Services.AddControllersWithViews(); // or AddRazorPages()

//var app = builder.Build();

//// Use the session middleware
//app.UseSession();

//app.MapControllers(); // Map your routes

//app.Run();
