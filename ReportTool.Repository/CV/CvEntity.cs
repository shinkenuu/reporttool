namespace ReportTool.Repository.CV
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class CvEntity : DbContext
    {
        public CvEntity()
            : base("name=CvEntity")
        {
        }

        public virtual DbSet<MONTHLY_CV> MONTHLY_CV { get; set; }
        public virtual DbSet<WEEKLY_CV> WEEKLY_CV { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
