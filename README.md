# ASP.NET-CORE-CRUD-OPEARTION

The Coffee Shop Management System schema outlined in your document has six tables: Product, User, Order, OrderDetail, Bills, and Customer, each with designated primary and foreign keys, relationships, and constraints.

# Overview of Each Table

# 1.Product Table:

Tracks products with attributes like ProductID (Primary Key), ProductName, ProductPrice, and Description.
Linked to the User table via UserID as a foreign key, potentially for tracking which user added or updated the product.

# 2.User Table:

Stores user account details such as UserID (Primary Key), UserName, Email, and IsActive.
Enables user-based authentication and role assignments within the application.

# 3.Order Table:

Records order details with attributes like OrderID (Primary Key), OrderDate, CustomerID, PaymentMode, and TotalAmount.
Has foreign keys CustomerID and UserID, linking it to the Customer and User tables respectively.

# 4.OrderDetail Table:

Contains individual items within each order, with OrderDetailID as the Primary Key.
Holds relationships with Order, Product, and User tables through foreign keys, enabling itemized billing.

# 5.Bills Table:

Tracks billing information with fields like BillID (Primary Key), BillNumber, BillDate, TotalAmount, and Discount.
Linked to Order and User tables through foreign keys, detailing payment and user responsibility.

# 6.Customer Table:

Manages customer data with CustomerID as the Primary Key and includes fields like CustomerName, Email, GST NO, and CityName.
Linked to User table through UserID, allowing users to manage customer records.


This schema ensures efficient tracking of coffee shop operations, enabling CRUD functionality across all essential entities for streamlined order management, billing, and user oversight
