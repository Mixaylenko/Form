using System.Security.Cryptography;
using System.Text;
using ServerForm.Interfaces;

namespace ServerForm.Services
{
    public class PasswordHasher : IPasswordHasher
    {
        public string HashPassword(string password)
        {
            using (var sha256 = SHA256.Create())
            {
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                var builder = new StringBuilder();

                foreach (var b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }

                return builder.ToString();
            }
        }

        public bool VerifyPassword(string password, string hashedPassword)
        {
            if (string.IsNullOrEmpty(password) || string.IsNullOrEmpty(hashedPassword))
                return false;

            return HashPassword(password) == hashedPassword;
        }
    }
}