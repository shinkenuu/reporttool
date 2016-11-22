namespace ReportTool.Repository.CV
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class MONTHLY_CV
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SAMPLE_DATE { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(80)]
        public string MAKE { get; set; }

        [Key]
        [Column(Order = 2)]
        [StringLength(80)]
        public string MODEL { get; set; }

        [Key]
        [Column(Order = 3)]
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

        public double? MSRP_PLUS_OPC { get; set; }

        public double? TP { get; set; }

        [Key]
        [Column(Order = 4)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int VOLUME { get; set; }
    }
}
