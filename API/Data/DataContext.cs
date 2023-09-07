using API.Models;
using Microsoft.EntityFrameworkCore;

namespace API.Data
{
    public class DataContext : DbContext
    {
        public DataContext(DbContextOptions options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasDefaultSchema("ActivitySchema");
            modelBuilder.Entity<ExcelData>()
                .ToTable("TimeSchedule", "ActivitySchema")
                .HasKey(row => row.Id);
            modelBuilder.Entity<ImportHistory>()
                .ToTable("ImportHistory", "ActivitySchema")
                .HasKey(file => file.Id);
        }

        public virtual DbSet<ExcelData>? TimeSchedule { get; set; }
        public virtual DbSet<ImportHistory>? ImportHistory { get; set; }
    }
}