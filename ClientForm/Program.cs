using Microsoft.AspNetCore.Authentication.Cookies;
using System.Net;
using Polly;
using Polly.Extensions.Http;
using System.Net.Security;
using System.Security.Authentication;
using System.Net.Http.Headers; // Для CacheControlHeaderValue
using Microsoft.AspNetCore.Http;
using ClientForm.Services; // Для SameSiteMode

var builder = WebApplication.CreateBuilder(args);

// 1. Базовая конфигурация
builder.Services.AddRazorPages(options =>
{
    options.Conventions.AuthorizeFolder("/");
    options.Conventions.AllowAnonymousToPage("/Auth/Login");
    options.Conventions.AllowAnonymousToPage("/Auth/Register");
    options.Conventions.AllowAnonymousToPage("/Error");
});

// 2. Конфигурация HttpClient с политиками устойчивости
var apiBaseUrl = builder.Configuration["ApiBaseUrl"]?.TrimEnd('/') + '/'
    ?? throw new InvalidOperationException("ApiBaseUrl не настроен в appsettings.json");

builder.Services.AddHttpClient("ServerAPI", client =>
{
    client.BaseAddress = new Uri(apiBaseUrl);
    client.DefaultRequestHeaders.Accept.Add(
        new MediaTypeWithQualityHeaderValue("application/json"));
    client.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue
    {
        NoCache = true
    };
    client.Timeout = TimeSpan.FromSeconds(30);
})
.ConfigurePrimaryHttpMessageHandler(() => new SocketsHttpHandler
{
    PooledConnectionLifetime = TimeSpan.FromMinutes(5),
    AutomaticDecompression = DecompressionMethods.All,
    SslOptions = new SslClientAuthenticationOptions
    {
        EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
        RemoteCertificateValidationCallback = (_, _, _, errors) =>
            builder.Environment.IsDevelopment() || errors == SslPolicyErrors.None
    }
})
.AddPolicyHandler(GetRetryPolicy())
.AddPolicyHandler(GetCircuitBreakerPolicy())
.AddPolicyHandler(Policy.TimeoutAsync<HttpResponseMessage>(TimeSpan.FromSeconds(10)));

// 3. Регистрация сервисов
builder.Services.AddScoped<IReportApiService, ReportApiService>();
builder.Services.AddHttpClient();
builder.Services.AddHttpContextAccessor();
builder.Services.AddAntiforgery(options =>
{
    options.HeaderName = "X-CSRF-TOKEN";
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
});

// 4. Настройка аутентификации
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/Auth/Login";
        options.AccessDeniedPath = "/AccessDenied";
        options.Cookie.Name = "Auth.Cookie";
        options.Cookie.HttpOnly = true;
        options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
        options.Cookie.SameSite = SameSiteMode.Lax;
        options.ExpireTimeSpan = TimeSpan.FromDays(1);
        options.SlidingExpiration = true;
        options.Events = new CookieAuthenticationEvents
        {
            OnRedirectToAccessDenied = context =>
            {
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                return Task.CompletedTask;
            }
        };
    });

// 5. Настройка сессии
builder.Services.AddSession(options =>
{
    options.Cookie.Name = "Session.Cookie";
    options.IdleTimeout = TimeSpan.FromMinutes(20);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.SameSite = SameSiteMode.Lax;
});

var app = builder.Build();

// 6. Конфигурация middleware pipeline
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers.Append(
            "Cache-Control", "public,max-age=604800");
    }
});

app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();
app.UseSession();

app.Use(async (context, next) =>
{
    context.Response.Headers.Append("X-Content-Type-Options", "nosniff");
    context.Response.Headers.Append("X-Frame-Options", "DENY");
});

app.MapRazorPages();
app.MapGet("/Report", () => Results.Redirect("/Report/Index"));
app.MapGet("/", () => Results.Redirect("/Index"));

app.Run();

// 7. Политики устойчивости
static IAsyncPolicy<HttpResponseMessage> GetRetryPolicy()
{
    return HttpPolicyExtensions
        .HandleTransientHttpError()
        .OrResult(msg => msg.StatusCode is
            HttpStatusCode.NotFound or
            HttpStatusCode.RequestTimeout or
            HttpStatusCode.GatewayTimeout)
        .WaitAndRetryAsync(3, retryAttempt =>
            TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
        onRetry: (outcome, delay, retryCount, _) =>
        {
            Console.WriteLine($"Retry {retryCount} after {delay.TotalSeconds}s due to: {outcome.Exception?.Message ?? outcome.Result.StatusCode.ToString()}");
        });
}

static IAsyncPolicy<HttpResponseMessage> GetCircuitBreakerPolicy()
{
    return HttpPolicyExtensions
        .HandleTransientHttpError()
        .CircuitBreakerAsync(
            handledEventsAllowedBeforeBreaking: 5,
            durationOfBreak: TimeSpan.FromSeconds(30),
            onBreak: (_, breakDelay) =>
                Console.WriteLine($"Circuit broken for {breakDelay.TotalSeconds}s"),
            onReset: () => Console.WriteLine("Circuit reset"),
            onHalfOpen: () => Console.WriteLine("Circuit half-open"));
}