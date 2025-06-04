using ServerForm.Models;

namespace ServerForm.Interfaces
{
    public interface IUserService
    {
        Task<UserModel> RegisterUserAsync(UserModel user);
        Task<UserModel> GetUserByIdAsync(int id);
        Task<UserModel> GetUserByEmailAsync(string Email);
        Task UpdateUserAsync(UserModel user);
        Task<UserModel> AuthenticateUserAsync(string email, string password); 
    }
}