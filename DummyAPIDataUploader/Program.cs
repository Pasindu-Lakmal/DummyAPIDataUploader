global using Microsoft.EntityFrameworkCore;
global using DummyAPIDataUploader.Models;
global using DummyAPIDataUploader.Data;
using DummyAPIDataUploader.Services.UploadLogsService;
using Syncfusion.Licensing;

var builder = WebApplication.CreateBuilder(args);



// Register Syncfusion license
SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UVhhQlVFfV1dXnxLflFyVWpTfll6dF1WACFaRnZdRl1iSXpRd0dgXH1ddXNS");
// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IUploadLogsService, UploadLogsService>();
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
