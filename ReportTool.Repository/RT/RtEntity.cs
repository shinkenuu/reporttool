namespace ReportTool.Repository.RT
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class RtEntity : DbContext
    {
        public RtEntity()
            : base("name=RtEntity")
        {
        }

        public virtual DbSet<MONTHLY_MSRP> MONTHLY_MSRP { get; set; }
        public virtual DbSet<MSRP_HISTORY> MSRP_HISTORY { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
