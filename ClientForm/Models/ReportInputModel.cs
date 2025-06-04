using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace ClientForm.Models
{
    public class ReportInputModel
    {
        [Required(ErrorMessage = "Название обязательно")]
        [StringLength(100, ErrorMessage = "Максимальная длина 100 символов")]
        public string Name { get; set; }

        [Required(ErrorMessage = "Файл обязателен")]
        public IFormFile File { get; set; }
    }
}