using Microsoft.EntityFrameworkCore;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
namespace mylib
{
    public class Phones
    {
        public int Id { get; set; }
        public string Product { get; set; }
        public string Company { get; set; }
        public int Price { get; set; }
        public int Count { get; set; }
    }
    public class AppConn : DbContext
    {
        public DbSet<Phones> phones { get; set; }
        public AppConn()
        {
            Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder builder)
        {
            builder.UseSqlite("Data Source = appdb.db");
        }
    }

    public class sqlwork
    {

        public void dbcreate()
        {
            using (AppConn db = new AppConn())
            {
                Phones model1 = new Phones { Id = 0, Product = "A51s", Company = "Samsung", Price = 20000, Count = 40 };
                Phones model2 = new Phones { Id = 1, Product = "5i", Company = "Realme", Price = 11000, Count = 150 };
                Phones model3 = new Phones { Id = 2, Product = "7 pro", Company = "Realme", Price = 17500, Count = 70 };
                Phones model4 = new Phones { Id = 3, Product = "3310", Company = "Nokia", Price = 1, Count = 1 };
                Phones model5 = new Phones { Id = 4, Product = "Note 7", Company = "Samsung", Price = 50000, Count = 100 };

                db.phones.Add(model1);
                db.phones.Add(model2);
                db.phones.Add(model3);
                db.phones.Add(model4);
                db.phones.Add(model5);
                db.SaveChanges();
            }
        }
    }

    public class ExcelWork
    {
        public void ExcelReader()
        {
            Application excel = new Application();
            if (excel == null)
            {
                throw "Excel is not installed";
            }

            Workbook wb = excel.Workbooks.Open(@"/tabledata.xlsx");
            Worksheet excelsheet = wb.Sheets[1];
            Range exRange = excelsheet.UsedRange;

            int rows = 
        }
    }
}