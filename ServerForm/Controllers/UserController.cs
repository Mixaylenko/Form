using Microsoft.AspNetCore.Mvc;
using ServerForm.Models;
using ServerForm.Interfaces;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authentication.Cookies;

namespace ServerForm.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class UserController : ControllerBase
    {
        private readonly IUserService _userService;
        private readonly DatabaseSettings _databaseSettings;

        public UserController(IUserService userService, IOptions<DatabaseSettings> databaseSettings)
        {
            _userService = userService;
            _databaseSettings = databaseSettings.Value;
        }

        // POST: api/user/register
        [HttpPost("register")]
        public async Task<IActionResult> Register([FromBody] RegisterRequest request)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var existingUser = await _userService.GetUserByEmailAsync(request.Email);
            if (existingUser != null)
            {
                return Conflict("Пользователь с таким email уже существует");
            }

            var newUser = new UserModel
            {
                Name = request.Name,
                Email = request.Email,
                Password = request.Password, // Пароль будет хеширован в сервисе
                Role = "user"
            };

            try
            {
                var createdUser = await _userService.RegisterUserAsync(newUser);
                return CreatedAtAction(nameof(GetUser), new { id = createdUser.Id }, new
                {
                    createdUser.Id,
                    createdUser.Name,
                    createdUser.Email,
                    createdUser.Role
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Ошибка при регистрации: {ex.Message}");
            }
        }

        // GET: api/user/{id}
        [HttpGet("{id}")]
        public async Task<IActionResult> GetUser(int id)
        {
            var user = await _userService.GetUserByIdAsync(id);
            if (user == null)
            {
                return NotFound();
            }
            return Ok(user);
        }

        // PUT: api/user/{id}
        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateUser(int id, [FromBody] UserModel updatedUser)
        {
            if (updatedUser == null || id != updatedUser.Id)
            {
                return BadRequest("Invalid user data.");
            }

            var existingUser = await _userService.GetUserByIdAsync(id);
            if (existingUser == null)
            {
                return NotFound();
            }

            // Обновление полей пользователя
            existingUser.Name = updatedUser.Name;
            existingUser.Email = updatedUser.Email;
            // Обновите другие поля по мере необходимости

            await _userService.UpdateUserAsync(existingUser);

            return NoContent();
        }


        // Авторизация
        [HttpPost("{UserId}")]
        public async Task<IActionResult> Login([FromBody] LoginRequest loginRequest)
        {
            if (loginRequest == null || string.IsNullOrEmpty(loginRequest.Email)
                || string.IsNullOrEmpty(loginRequest.Password))
            {
                return BadRequest("Email and Password are required.");
            }

            var user = await _userService.AuthenticateUserAsync(loginRequest.Email, loginRequest.Password);

            if (user == null)
                return Unauthorized("Неверный email или пароль");

                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.NameIdentifier, user.Id.ToString()),
                    new Claim(ClaimTypes.Name, user.Name.Trim()), // Важно: сохраняем очищенное имя
                    new Claim(ClaimTypes.Email, user.Email),
                    new Claim(ClaimTypes.Role, user.Role ?? "user")
                };

                var claimsIdentity = new ClaimsIdentity(
                claims, CookieAuthenticationDefaults.AuthenticationScheme);

            var authProperties = new AuthenticationProperties
            {
                IsPersistent = true,
                ExpiresUtc = DateTimeOffset.UtcNow.AddDays(30),
                AllowRefresh = true
            };

            await HttpContext.SignInAsync(
                CookieAuthenticationDefaults.AuthenticationScheme,
                new ClaimsPrincipal(claimsIdentity),
                authProperties);

            return Ok(new
            {
                Message = "Login successful",
                UserId = user.Id,
                Name = user.Name,
                Role = user.Role ?? "user"
            });
        }
    }

    //DTO:
    public class LoginRequest
    {
        public string Email { get; set; }
        public string Password { get; set; }
    }
    public class RegisterRequest
    {
        [Required]
        public string Name { get; set; }

        [Required]
        [EmailAddress]
        public string Email { get; set; }

        [Required]
        [MinLength(6)]
        public string Password { get; set; }
    }
}