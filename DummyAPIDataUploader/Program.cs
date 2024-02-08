global using Microsoft.EntityFrameworkCore;
global using DummyAPIDataUploader.Models;
global using DummyAPIDataUploader.Data;
using DummyAPIDataUploader.Services.UploadLogsService;
using Syncfusion.Licensing;
using DummyAPIDataUploader.Services.ExcelManipulationServices;

var builder = WebApplication.CreateBuilder(args);



// Register Syncfusion license
SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NAaF1cXmhKYVRpR2Nbe05yflFOal9UVAciSV9jS3pTdEVlWX1aeHdTQWFfVg==");
// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IUploadLogsService, UploadLogsService>();
builder.Services.AddScoped<IExcelManipulationService, ExcelManipulationService>();
builder.Services.AddDbContext<DataContext>();
builder.Services.AddCors(option => option.AddPolicy(name:"uploader" , 
    
    policy=>
    {
        policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader();
    }
    
    ));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("uploader");

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
