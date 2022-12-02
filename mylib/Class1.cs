using Microsoft.EntityFrameworkCore;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore.Update.Internal;
using Microsoft.EntityFrameworkCore.Storage;

namespace mylib
{
    
    public class Phone
    {
        public int Id { get; set; }
        public string Product { get; set; }
        public string Company { get; set; }
        public int Price { get; set; }
        public int Count { get; set; }
    }
    public class AppConn : DbContext
    {
        public bool iscreated = false;
        public DbSet<Phone> Phones => Set<Phone>();
        public  AppConn()
        {
            if (Database.EnsureCreated() == true)
            {
                iscreated = true;
            }
        }

        protected override void OnConfiguring(DbContextOptionsBuilder builder)
        {
            builder.UseSqlite("Data Source = appdb.db");
            builder.EnableSensitiveDataLogging();
        }
    }

    public class sqlwork
    {
        public void Dbcreate()
        {
           using (AppConn db = new AppConn())
           {
                if (db.iscreated == true)
                {
                    Phone model1 = new Phone { Product = "A51s", Company = "Samsung", Price = 20000, Count = 40 };
                    Phone model2 = new Phone { Product = "5i", Company = "Realme", Price = 11000, Count = 150 };
                    Phone model3 = new Phone { Product = "7 pro", Company = "Realme", Price = 17500, Count = 70 };
                    Phone model4 = new Phone { Product = "3310", Company = "Nokia", Price = 1, Count = 1 };
                    Phone model5 = new Phone { Product = "Note 7", Company = "Samsung", Price = 50000, Count = 100 };

                    db.Phones.Add(model1);
                    db.Phones.Add(model2);
                    db.Phones.Add(model3);
                    db.Phones.Add(model4);
                    db.Phones.Add(model5);
                    db.SaveChanges();
                }
           }
        }

        public List<Phone> DbShowInfo()
        {
            AppConn db = new AppConn();
            var result = db.Phones.ToList();
            return result;
        }
    }

    public class ExcelWork
    {
        public void ExcelReader()
        {
            var ExApp = new Excel.Application();
            if (ExApp == null)
            {
                //TODO:Кинуть сюда экспешн
            }

            Excel.Workbook wb = ExApp.Workbooks.Open(@"/tabledata.xlsx");
            Excel._Worksheet excelsheet = wb.Sheets[1];
            Excel.Range exRange = excelsheet.UsedRange;

            //int rows = 
        }



        public void ExcelWriter()
        {
            var ExApp = new Excel.Application();
            if (ExApp == null)
            {
                //TODO: Кинуть сюда экспешн
            }

            

            ExApp.Visible = true;

            ExApp.Workbooks.Add();

            Excel._Worksheet worksheet = (Excel.Worksheet)ExApp.ActiveSheet;

            using (AppConn db = new AppConn())
            {
               var res = db.Phones.ToList();
                int row = 1;


                worksheet.Cells[1, "A"] = "ID";
                worksheet.Cells[1, "B"] = "Product";
                worksheet.Cells[1, "C"] = "Company";
                worksheet.Cells[1, "D"] = "Price";
                worksheet.Cells[1, "E"] = "Count";

                foreach (Phone item in res)
                {
                    row++;
                    worksheet.Cells[row, "A"] = item.Id;
                    worksheet.Cells[row, "B"] = item.Product;
                    worksheet.Cells[row, "C"] = item.Company;
                    worksheet.Cells[row, "D"] = item.Price;
                    worksheet.Cells[row, "E"] = item.Count;
                }
            }
        }
    }
}