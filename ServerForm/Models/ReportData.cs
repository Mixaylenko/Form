using System.ComponentModel.DataAnnotations;

namespace ServerForm.Models
{
    public class ReportData
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [StringLength(100)]
        public string FileName { get; set; }
        public string FilePath { get; set; }

        public string Name { get; set; }
    }
}
