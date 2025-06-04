using Microsoft.EntityFrameworkCore;

namespace ServerForm.Models
{
    public class DatabaseContext : DbContext
    {
        public DatabaseContext(DbContextOptions<DatabaseContext> options)
            : base(options)
        {
        }

        public DbSet<UserModel> Users { get; set; }
        public DbSet<ReportData> ReportDatas { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
        }
    }
}