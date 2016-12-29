namespace ReportTool.Repository.RT
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class MSRP_HISTORY
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DATADATE { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int UID { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(80)]
        public string MAKE { get; set; }

        [Key]
        [Column(Order = 3)]
        [StringLength(80)]
        public string MODEL { get; set; }

        [Key]
        [Column(Order = 4)]
        [StringLength(80)]
        public string VERSION { get; set; }

        [StringLength(8)]
        public string PY { get; set; }

        [StringLength(8)]
        public string MY { get; set; }

        [StringLength(2)]
        public string DOORS { get; set; }

        [StringLength(15)]
        public string BODY_TYPE { get; set; }

        public double? MSRP { get; set; }
    }
}
