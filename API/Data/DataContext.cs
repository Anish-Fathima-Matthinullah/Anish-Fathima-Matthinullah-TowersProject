using API.Modals;
using Microsoft.EntityFrameworkCore;

namespace API.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions options) : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasDefaultSchema("ActivitySchema");
            modelBuilder.Entity<ExcelData>()
                .ToTable("Buildings", "ActivitySchema")
                .HasKey(row => row.Id);
            modelBuilder.Entity<ImportHistory>()
                .ToTable("ImportHistory", "ActivitySchema")
                .HasKey(file => file.Id);
                modelBuilder.Entity<Header>()
                .ToTable("Header", "ActivitySchema")
                .HasKey(h => h.Id);
            modelBuilder.Entity<Activity>()
                .ToTable("Activity", "ActivitySchema")
                .HasKey(a => a.Id);
        }

        public virtual DbSet<ExcelData>? Buildings { get; set; }
        public virtual DbSet<ImportHistory>? ImportHistory { get; set; }
        public virtual DbSet<Header>? Header { get; set; }
        public virtual DbSet<Activity>? Activity { get; set; }
    }
}