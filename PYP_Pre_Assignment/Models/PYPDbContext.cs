using Microsoft.EntityFrameworkCore;
using PYP_Pre_Assignment.Models;
namespace PYP_Pre_Assignment.Models
{
    public class PYPDbContext : DbContext
    {
        public PYPDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<XLSFile> XLSFiles { get; set; }
    }
}
