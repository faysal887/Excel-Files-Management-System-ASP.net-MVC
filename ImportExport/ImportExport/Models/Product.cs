using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ImportExport.Models
{
    public class Product
    {
        //Product table content
        public int Id  //record id
        {
            get;
            set;
        }

        public string FirstName
        {
            get;
            set;
        }
        public string LastName
        {
            get;
            set;
        }
        public string Email
        {
            get;
            set;
        }
        public string Title
        {
            get;
            set;
        }
        public string Address1
        {
            get;
            set;
        }
        public string Address2
        {
            get;
            set;
        }
        public string City
        {
            get;
            set;
        }
        public string COAddress
        {
            get;
            set;
        }
        public string Zip
        {
            get;
            set;
        }
        public string CompanyEmail
        {
            get;
            set;
        }
        public string CompanyName
        {
            get;
            set;
        }
        public string CompanyLegalName
        {
            get;
            set;
        }
        public string Department
        {
            get;
            set;
        }
        public string CompanyPhone
        {
            get;
            set;
        }
        public string EmployeesMin
        {
            get;
            set;
        }
        public string EmployeesMax
        {
            get;
            set;
        }
        public string CompanyOrganisationNumber
        {
            get;
            set;
        }
        public string VatNumber
        {
            get;
            set;
        }
        public int TurnOver
        {
            get;
            set;
        }
        public string County
        {
            get;
            set;
        }
        public string Minicipal
        {
            get;
            set;
        }
        public string SniCode
        {
            get;
            set;
        }
        public string Branch
        {
            get;
            set;
        }
        public string CompanyType
        {
            get;
            set;
        }
        public string Sector
        {
            get;
            set;
        }
        public string MasterSniCode
        {
            get;
            set;
        }
        public string MasterBranch
        {
            get;
            set;
        }

        //Product table content

        public int fid // excel file id
        {
            get;
            set;
        }

        public DateTime TimeStamp
        {
            get;
            set;
        }

        public int NoOfPosts
        {
            get;
            set;
        }

        public string FileName
        {
            get;
            set;
        }




        // public static List<Product> ExcelData = new List<Product>();

        public static void addExcelFileDataToDB(List<Product> ExcelFileData)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {

                // Step 1: Connecting to Database
                connection.Open();

                // Step 2: Creating Command
                SqlCommand command = connection.CreateCommand();

                // "INSERT INTO Contacts (Name, Contact, Email) VALUES ('Usman', '03312873', 'usman@gmail.com')";
                foreach (var item in ExcelFileData)
                {
                    string insertQuery = "INSERT INTO Product (Id, firstname, lastname, email, title, address1, address2, city, coaddress, zip, companyemail, companyname, companylegalname, department, companyphone, employeesmin, employeesmax, companyorganisationnumber, vatnumber, turnover, county, minicipal, snicode, branch, companytype, sector, mastersnicode, masterbranch, excelfilename) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', '{22}', '{23}', '{24}', '{25}', '{26}', '{27}', '{28}' )";
                    insertQuery = String.Format(insertQuery, item.Id, item.FirstName, item.LastName, item.Email, item.Title, item.Address1, item.Address2, item.City, item.COAddress, item.Zip, item.CompanyEmail, item.CompanyName, item.CompanyLegalName, item.Department, item.CompanyPhone, item.EmployeesMin, item.EmployeesMax, item.CompanyOrganisationNumber, item.VatNumber, item.TurnOver, item.County, item.Minicipal, item.SniCode, item.Branch, item.CompanyType, item.Sector, item.MasterSniCode, item.MasterBranch, item.FileName);

                    command.CommandText = insertQuery;
                    command.Connection = connection;

                    // Step 3: Read Data from Table
                    var reader = command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }

        public static void addExcelFileNameToDB(List<Product> ExcelFileName)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {

                // Step 1: Connecting to Database
                connection.Open();

                // Step 2: Creating Command
                SqlCommand command = connection.CreateCommand();

                // "INSERT INTO Contacts (Name, Contact, Email) VALUES ('Usman', '03312873', 'usman@gmail.com')";

                string insertQuery = "INSERT INTO ExcelFiles (name, datetime, noofposts) VALUES ('{0}', '{1}', '{2}' )";
                insertQuery = String.Format(insertQuery, ExcelFileName[0].FileName, ExcelFileName[0].TimeStamp, ExcelFileName[0].NoOfPosts);

                command.CommandText = insertQuery;
                command.Connection = connection;

                // Step 3: Read Data from Table
                var reader = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }


        public static void DeleteExcelFile(string fn)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();

                //command.CommandText = "DELETE FROM EXCELFILES WHERE name LIKE '%" + fn + "%'";
                command.CommandText = "DELETE FROM EXCELFILES WHERE name = '" + fn + "'";
                command.Connection = connection;

                var reader = command.ExecuteNonQuery();



            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }


        public async static Task<List<Product>> getExcelFileNames()
        {
            // Get phonebook from DB

            List<Product> list = new List<Product>();
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {

                // Step 1: Connecting to Database
                connection.Open();

                // Step 2: Creating Command
                SqlCommand command = connection.CreateCommand();

                command.CommandText = "SELECT * FROM ExcelFiles";
                command.Connection = connection;

                // Step 3: Read Data from Table
                var reader = command.ExecuteReader();

                Product p;

                while (reader.Read())
                {
                    p = new Product();
                    p.fid = (int)reader["Id"];
                    p.FileName = (string)reader["name"];
                    p.TimeStamp = (DateTime)reader["datetime"];
                    p.NoOfPosts = (int)reader["noofposts"];
                    list.Add(p);
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
            return list;
        }


        internal static void DeleteExcels()
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();

                command.CommandText = "Delete ExcelFiles";
                command.Connection = connection;

                var reader = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }


        public async static Task<List<Product>> ShowExcelFile(string fn)
        {

            List<Product> list = new List<Product>();
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            try
            {

                // Step 1: Connecting to Database
                connection.Open();

                // Step 2: Creating Command
                SqlCommand command = connection.CreateCommand();

                string query = "SELECT * FROM Product WHERE excelfilename = '" + fn + "'";
                //  "SELECT * FROM Products WHERE isDeleted=0 AND name LIKE '%" + name + "%'" + "OR category LIKE '%" + name + "%'";

                //  string query = "SELECT * FROM Product WHERE excelfilename=" + "file1" ;
                command.CommandText = query;
                command.Connection = connection;

                // Step 3: Read Data from Table
                var reader = command.ExecuteReader();

                Product p;

                while (reader.Read())
                {
                    p = new Product();
                    p.Id = (int)reader["Id"];
                    p.FirstName = (string)reader["firstname"];
                    p.LastName = (string)reader["lastname"];
                    p.Email = (string)reader["email"];
                    p.Title = (string)reader["title"];
                    p.Address1 = (string)reader["address1"];
                    p.Address2 = (string)reader["address2"];
                    p.City = (string)reader["city"];
                    p.COAddress = (string)reader["coaddress"];
                    p.Zip = (string)reader["zip"];
                    p.CompanyEmail = (string)reader["companyemail"];
                    p.CompanyName = (string)reader["companyname"];
                    p.CompanyLegalName = (string)reader["companylegalname"];
                    p.Department = (string)reader["department"];
                    p.CompanyPhone = (string)reader["companyphone"];
                    p.EmployeesMin = (string)reader["employeesmin"];
                    p.EmployeesMax = (string)reader["employeesmax"];
                    p.CompanyOrganisationNumber = (string)reader["companyorganisationnumber"];
                    p.VatNumber = (string)reader["vatnumber"];
                    p.TurnOver = (int)reader["turnover"];
                    p.County = (string)reader["county"];
                    p.Minicipal = (string)reader["minicipal"];
                    p.SniCode = (string)reader["snicode"];
                    p.Branch = (string)reader["branch"];
                    p.CompanyType = (string)reader["companytype"];
                    p.Sector = (string)reader["sector"];
                    p.MasterSniCode = (string)reader["mastersniCode"];
                    p.MasterBranch = (string)reader["masterbranch"];
                    list.Add(p);
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }

            return list;
        }


        public async static Task<List<Product>> searchByTurnover(int val, int temp)
        {

            List<Product> list = new List<Product>();
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

            try
            {

                // Step 1: Connecting to Database
                connection.Open();

                // Step 2: Creating Command
                if (temp == 1)
                {
                    SqlCommand command = connection.CreateCommand();

                    string query = "SELECT * FROM Product WHERE turnover > " + val;
                    //  "SELECT * FROM Products WHERE isDeleted=0 AND name LIKE '%" + name + "%'" + "OR category LIKE '%" + name + "%'";

                    //  string query = "SELECT * FROM Product WHERE excelfilename=" + "file1" ;
                    command.CommandText = query;
                    command.Connection = connection;

                    // Step 3: Read Data from Table
                    var reader = command.ExecuteReader();

                    Product p;

                    while (reader.Read())
                    {
                        p = new Product();
                        p.Id = (int)reader["Id"];
                        p.FirstName = (string)reader["firstname"];
                        p.LastName = (string)reader["lastname"];
                        p.Email = (string)reader["email"];
                        p.Title = (string)reader["title"];
                        p.Address1 = (string)reader["address1"];
                        p.Address2 = (string)reader["address2"];
                        p.City = (string)reader["city"];
                        p.COAddress = (string)reader["coaddress"];
                        p.Zip = (string)reader["zip"];
                        p.CompanyEmail = (string)reader["companyemail"];
                        p.CompanyName = (string)reader["companyname"];
                        p.CompanyLegalName = (string)reader["companylegalname"];
                        p.Department = (string)reader["department"];
                        p.CompanyPhone = (string)reader["companyphone"];
                        p.EmployeesMin = (string)reader["employeesmin"];
                        p.EmployeesMax = (string)reader["employeesmax"];
                        p.CompanyOrganisationNumber = (string)reader["companyorganisationnumber"];
                        p.VatNumber = (string)reader["vatnumber"];
                        p.TurnOver = (int)reader["turnover"];
                        p.County = (string)reader["county"];
                        p.Minicipal = (string)reader["minicipal"];
                        p.SniCode = (string)reader["snicode"];
                        p.Branch = (string)reader["branch"];
                        p.CompanyType = (string)reader["companytype"];
                        p.Sector = (string)reader["sector"];
                        p.MasterSniCode = (string)reader["mastersniCode"];
                        p.MasterBranch = (string)reader["masterbranch"];
                        list.Add(p);
                    }
                }

                else if (temp == 0)
                {
                    SqlCommand command = connection.CreateCommand();

                    string query = "SELECT * FROM Product WHERE turnover < " + val;
                    //  "SELECT * FROM Products WHERE isDeleted=0 AND name LIKE '%" + name + "%'" + "OR category LIKE '%" + name + "%'";

                    //  string query = "SELECT * FROM Product WHERE excelfilename=" + "file1" ;
                    command.CommandText = query;
                    command.Connection = connection;

                    // Step 3: Read Data from Table
                    var reader = command.ExecuteReader();

                    Product p;

                    while (reader.Read())
                    {
                        p = new Product();
                        p.Id = (int)reader["Id"];
                        p.FirstName = (string)reader["firstname"];
                        p.LastName = (string)reader["lastname"];
                        p.Email = (string)reader["email"];
                        p.Title = (string)reader["title"];
                        p.Address1 = (string)reader["address1"];
                        p.Address2 = (string)reader["address2"];
                        p.City = (string)reader["city"];
                        p.COAddress = (string)reader["coaddress"];
                        p.Zip = (string)reader["zip"];
                        p.CompanyEmail = (string)reader["companyemail"];
                        p.CompanyName = (string)reader["companyname"];
                        p.CompanyLegalName = (string)reader["companylegalname"];
                        p.Department = (string)reader["department"];
                        p.CompanyPhone = (string)reader["companyphone"];
                        p.EmployeesMin = (string)reader["employeesmin"];
                        p.EmployeesMax = (string)reader["employeesmax"];
                        p.CompanyOrganisationNumber = (string)reader["companyorganisationnumber"];
                        p.VatNumber = (string)reader["vatnumber"];
                        p.TurnOver = (int)reader["turnover"];
                        p.County = (string)reader["county"];
                        p.Minicipal = (string)reader["minicipal"];
                        p.SniCode = (string)reader["snicode"];
                        p.Branch = (string)reader["branch"];
                        p.CompanyType = (string)reader["companytype"];
                        p.Sector = (string)reader["sector"];
                        p.MasterSniCode = (string)reader["mastersniCode"];
                        p.MasterBranch = (string)reader["masterbranch"];
                        list.Add(p);
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }

            return list;
        }


        internal static void DeleteMultipleProducts(string[] pids)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();

                foreach (var pid in pids)
                {
                    command.CommandText = "DELETE from Product WHERE Id = '" + pid + "'";
                    command.Connection = connection;

                    var reader = command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }

        public static void UpdateNoOfPosts(int count)
        {
            SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
            try
            {
                connection.Open();

                SqlCommand command = connection.CreateCommand();

                command.CommandText = "UPDATE ExcelFiles SET noofposts= " + count;
                command.Connection = connection;

                var reader = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            { }
            finally
            {
                connection.Close();
            }
        }

        /*  public async static Task<List<Product>> GetColumnNamesOfProductTable()
          {
              // Get phonebook from DB

              List<Product> list = new List<Product>();
              SqlConnection connection = new SqlConnection(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=Products;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
              try
              {

                  // Step 1: Connecting to Database
                  connection.Open();

                  // Step 2: Creating Command
                  SqlCommand command = connection.CreateCommand();



                  command.CommandText = "SELECT * FROM PRODUCTS.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME= "+ "PRODUCT";
                  command.Connection = connection;

                  // Step 3: Read Data from Table
                  var reader = command.ExecuteReader();

                  Product p;

                  while (reader.Read())
                  {
                      p = new Product();
                      p.fid = (int)reader["Id"];
                      p.FileName = (string)reader["name"];
                      p.TimeStamp = (DateTime)reader["datetime"];
                      p.NoOfPosts = (int)reader["noofposts"];
                      list.Add(p);
                  }
              }
              catch (Exception ex)
              { }
              finally
              {
                  connection.Close();
              }
              return list;
          }*/


    }
};