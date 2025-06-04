using System.ComponentModel.DataAnnotations;

namespace ServerForm.Models
{
    public class UserModel
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [StringLength(100)]
        public string Name { get; set; }

        [Required]
        [EmailAddress]
        public string Email { get; set; }

        [Required]
        [StringLength(100)]
        public string Password { get; set; }
        [Required]
        [StringLength(100)]
        public string Role { get; set; }

        public DateTime RegisteredAt { get; set; } = DateTime.UtcNow;
    }
}
