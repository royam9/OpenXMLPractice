using Microsoft.OpenApi.Models;
using Services;
using Services.Interfaces;
using static Services.WordService;

namespace OpenXMLPractice
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            #region Swagger
            builder.Services.AddSwaggerGen(swagger =>
            {
                //�ҥ�SwaggerResponse����
                //swagger.EnableAnnotations();

                //swagger.SwaggerDoc("v1", new OpenApiInfo
                //{
                //    Title = "Api",
                //    Version = "v1",
                //    Description = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Swagger", "SwaggerDescription.html"))
                //});

                // To Enable authorization using Swagger (JWT)    
                swagger.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme()
                {
                    Name = "Authorization",
                    Type = SecuritySchemeType.ApiKey,
                    Scheme = "Bearer",
                    BearerFormat = "JWT",
                    In = ParameterLocation.Header,
                    Description = "Enter 'Bearer' [space] and then your valid token in the text input below.\r\n\r\nExample: \"Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9\"",
                });

                swagger.AddSecurityRequirement(new OpenApiSecurityRequirement
                {
                    {
                        new OpenApiSecurityScheme
                        {
                            Reference = new OpenApiReference
                            {
                                Type = ReferenceType.SecurityScheme,
                                Id = "Bearer"
                            }
                        },
                        Array.Empty<string>()
                    }
                });

                // �W�[ data �y�z
                //var apiXmlPath = Path.Combine(AppContext.BaseDirectory, "Api.xml");
                //var modelXmlPath = Path.Combine(AppContext.BaseDirectory, "Models.xml");
                //swagger.IncludeXmlComments(apiXmlPath);
                //swagger.IncludeXmlComments(modelXmlPath);
            });
            #endregion

            builder.Services.AddScoped<IExcelService, ExcelService>();
            builder.Services.AddScoped<IWordService, CSDNSolution>();

            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();


            app.MapControllers();

            app.Run();
        }
    }
}
