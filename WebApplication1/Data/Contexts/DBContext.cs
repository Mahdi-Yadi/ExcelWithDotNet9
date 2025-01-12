using ExcelWithDotNet9.Web.Data.entities;
using Microsoft.EntityFrameworkCore;
namespace ExcelWithDotNet9.Web.Data.Contexts;
public class DBContext : DbContext
{
    public DBContext(DbContextOptions<DBContext> options): base(options)
    {
        
    }
    public DbSet<Order> Order { get; set; }

}