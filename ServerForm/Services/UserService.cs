using Microsoft.EntityFrameworkCore;
using ServerForm.Interfaces;
using ServerForm.Models;
using System.Security.Cryptography;
using System.Text;

namespace ServerForm.Services
{
    public class UserService : IUserService
    {
        private readonly DatabaseContext _context;
        private readonly IPasswordHasher _passwordHasher;

        public UserService(DatabaseContext context, IPasswordHasher passwordHasher)
        {
            _context = context;
            _passwordHasher = passwordHasher;
        }

        public async Task<UserModel> GetUserByIdAsync(int id)
        {
            return await _context.Users.AsNoTracking().FirstOrDefaultAsync(u => u.Id == id);
        }

        public async Task<UserModel> GetUserByEmailAsync(string email)
        {
            return await _context.Users.AsNoTracking().FirstOrDefaultAsync(u => u.Email == email);
        }

        public async Task UpdateUserAsync(UserModel user)
        {
            var existingUser = await _context.Users.FindAsync(user.Id);
            if (existingUser == null)
                throw new ArgumentException("User not found");

            existingUser.Name = user.Name;
            existingUser.Email = user.Email;
            existingUser.Role = user.Role;
            
            if (!string.IsNullOrEmpty(user.Password))
            {
                existingUser.Password = _passwordHasher.HashPassword(user.Password);
            }

            await _context.SaveChangesAsync();
        }

        public async Task<UserModel> RegisterUserAsync(UserModel user)
        {
            if (user == null)
                throw new ArgumentNullException(nameof(user));

            if (string.IsNullOrWhiteSpace(user.Password))
                throw new ArgumentException("Password is required");

            if (await _context.Users.AnyAsync(u => u.Email == user.Email))
                throw new InvalidOperationException("Email already exists");

            user.Password = _passwordHasher.HashPassword(user.Password);
            user.RegisteredAt = DateTime.UtcNow;
            user.Role ??= "user"; // Устанавливаем роль по умолчанию

            _context.Users.Add(user);
            await _context.SaveChangesAsync();

            return user;
        }

        public async Task<UserModel> AuthenticateUserAsync(string email, string password)
        {
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(password))
                return null;

            var user = await _context.Users
                .AsNoTracking()
                .FirstOrDefaultAsync(u => u.Email == email);

            if (user == null)
                return null;

            if (!_passwordHasher.VerifyPassword(password, user.Password))
                return null;

            return user;
        }
    }
}