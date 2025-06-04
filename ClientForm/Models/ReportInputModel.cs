// Models/ReportInputModel.cs
using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace ClientForm.Models
{
    public class ReportInputModel
    {
        [Required]
        public string Name { get; set; }

        public IFormFile? File { get; set; }
    }
}