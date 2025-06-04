using Microsoft.OpenApi.Models;
using ServerForm.Interfaces;
using ServerForm.Services;
using ServerForm.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Authentication.Cookies;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Reflection;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.StaticFiles;

var builder = WebApplication.CreateBuilder(args);

// Добавляем контроллеры и настройки Swagger
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "Report System API",
        Version = "v1",
        Description = "API for managing reports and users",
        Contact = new OpenApiContact
        {
            Name = "Support",
            Email = "support@reportsystem.com"
        }
    });

    // Добавляем поддержку аутентификации в Swagger
    c.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Description = "JWT Authorization header using the Bearer scheme",
        Type = SecuritySchemeType.Http,
        Scheme = "bearer"
    });

    // Добавляем поддержку загрузки файлов
    c.OperationFilter<FileUploadOperationFilter>();

    // Включаем XML комментарии
    var xmlFile = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
    {
        c.IncludeXmlComments(xmlPath);
    }
});

// Настройка CORS для клиента
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowClient", policy =>
    {
        policy.WithOrigins("https://localhost:7012", "http://localhost:5110")
              .AllowAnyHeader()
              .AllowAnyMethod()
              .AllowCredentials();
    });
});

// Привязка настроек базы данных
builder.Services.Configure<DatabaseSettings>(builder.Configuration.GetSection("DatabaseSettings"));

builder.Services.AddSingleton<IContentTypeProvider, FileExtensionContentTypeProvider>();
// Регистрация сервисов
builder.Services.AddSingleton<IContentTypeProvider, FileExtensionContentTypeProvider>();
builder.Services.AddScoped<IReportService, ReportService>();
builder.Services.AddScoped<IUserService, UserService>();
builder.Services.AddScoped<IPasswordHasher, PasswordHasher>();

// Настройка DbContext с SQL Server
builder.Services.AddDbContext<DatabaseContext>((serviceProvider, options) =>
{
    var dbSettings = serviceProvider.GetRequiredService<IOptions<DatabaseSettings>>().Value;
    options.UseSqlServer(dbSettings.ConnectionParameters.ConnectionString);
}, ServiceLifetime.Scoped);

// Настройка аутентификации
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.Cookie.Name = "AuthCookie";
        options.Cookie.HttpOnly = true;
        options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
        options.Cookie.SameSite = SameSiteMode.None;
        options.SlidingExpiration = true;
        options.ExpireTimeSpan = TimeSpan.FromDays(30);
        options.LoginPath = "/api/User/login";
        options.LogoutPath = "/api/User/logout";
        options.AccessDeniedPath = "/api/User/accessdenied";
    });

var app = builder.Build();

// Конфигурация middleware pipeline
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

// Настройка Swagger UI
app.UseSwagger();
app.UseSwaggerUI(c =>
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Report System API v1");
    c.RoutePrefix = "swagger";
    c.DisplayRequestDuration();
});

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

// Важно правильное расположение этих middleware
app.UseCors("AllowClient");
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

// Редирект с корня на Swagger UI
app.MapGet("/", () => Results.Redirect("/swagger"));

app.Run();

// Фильтр для обработки загрузки файлов в Swagger
public class FileUploadOperationFilter : IOperationFilter
{
    public void Apply(OpenApiOperation operation, OperationFilterContext context)
    {
        var formParameters = context.ApiDescription.ActionDescriptor.Parameters
            .Where(x => x.BindingInfo?.BindingSource?.Id == "Form")
            .SelectMany(x => x.ParameterType.GetProperties())
            .ToList();

        if (formParameters.Any())
        {
            operation.RequestBody = new OpenApiRequestBody
            {
                Content =
                {
                    ["multipart/form-data"] = new OpenApiMediaType
                    {
                        Schema = new OpenApiSchema
                        {
                            Type = "object",
                            Properties = formParameters.ToDictionary(
                                x => x.Name,
                                x => new OpenApiSchema
                                {
                                    Type = x.PropertyType == typeof(IFormFile) ? "string" : "string",
                                    Format = x.PropertyType == typeof(IFormFile) ? "binary" : null
                                })
                        }
                    }
                }
            };
        }
    }
}