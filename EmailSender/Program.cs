using Microsoft.AspNetCore.Http.Features;
using Microsoft.OpenApi.Models;

var builder = WebApplication.CreateBuilder(args);

// Bind Graph options from configuration
builder.Services.Configure<EmailSender.Services.GraphOptions>(builder.Configuration.GetSection("Graph"));

// Register email sender service
builder.Services.AddScoped<EmailSender.Services.IEmailSender, EmailSender.Services.GraphEmailSender>();

// Increase request body limits for large attachments
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = null; // unlimited
});

// Add services to the container.
// Configure OpenAPI/Swagger
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo { Title = "EmailSender API", Version = "v1" });
});

builder.Services.AddControllers();

// Increase multipart/form-data limits globally
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = long.MaxValue;
    options.ValueLengthLimit = int.MaxValue;
    options.MemoryBufferThreshold = int.MaxValue;
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "EmailSender v1");
        c.RoutePrefix = string.Empty; // serve UI at app root
    });
}

app.UseHttpsRedirection();

app.MapControllers();

app.Run();
