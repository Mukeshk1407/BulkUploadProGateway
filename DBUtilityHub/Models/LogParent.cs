﻿using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DBUtilityHub.Models
{
    public class LogParent
    {
        [Key]
        public int ID { get; set; }
        public DateTime Timestamp { get; set; }
        public int RecordCount { get; set; }
        public int PassCount { get; set; }
        public int FailCount { get; set; }
        public int User_Id { get; set; }
        public string FileName { get; set; }

        [ForeignKey("Entity_Id")]
        public int Entity_Id { get; set; }
    }
}
