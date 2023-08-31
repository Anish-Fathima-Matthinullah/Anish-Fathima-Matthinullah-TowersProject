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
            modelBuilder.Entity<Building>()
                .ToTable("Building", "ActivitySchema")
                .HasKey(b => b.Id);
            modelBuilder.Entity<Tower>()
                .ToTable("Tower", "ActivitySchema")
                .HasKey(t => t.Id);
            modelBuilder.Entity<Milestone>()
                .ToTable("Milestone", "ActivitySchema")
                .HasKey(m => m.Id);
            modelBuilder.Entity<Area>()
                .ToTable("Area", "ActivitySchema")
                .HasKey(a => a.Id);
            modelBuilder.Entity<Region>()
                .ToTable("Region", "ActivitySchema")
                .HasKey(r => r.Id);
            modelBuilder.Entity<Floor>()
                .ToTable("Floor", "ActivitySchema")
                .HasKey(f => f.Id);
            modelBuilder.Entity<Work>()
                .ToTable("Work", "ActivitySchema")
                .HasKey(w => w.Id);
            modelBuilder.Entity<Activity>()
                .ToTable("Activity", "ActivitySchema")
                .HasKey(a => a.Id);
        }

        public virtual DbSet<ExcelData>? Buildings { get; set; }
        public virtual DbSet<ImportHistory>? ImportHistory { get; set; }
        public virtual DbSet<Building>? Building { get; set; }
        public virtual DbSet<Tower>? Tower { get; set; }
        public virtual DbSet<Milestone>? Milestone { get; set; }
        public virtual DbSet<Area>? Area { get; set; }
        public virtual DbSet<Region>? Region { get; set; }
        public virtual DbSet<Floor>? Floor { get; set; }
        public virtual DbSet<Work>? Work { get; set; }
        public virtual DbSet<Activity>? Activity { get; set; }
    }
}