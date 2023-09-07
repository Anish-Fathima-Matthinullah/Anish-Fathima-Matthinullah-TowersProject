using API.Data;
using API.Interfaces;
using API.Services;
using Microsoft.EntityFrameworkCore;

namespace API.Extensions
{
    public static class ApplicationServiceExtension
    {
        public static IServiceCollection AddApplicationServices
        (this IServiceCollection services, IConfiguration config)
        {
            services.AddDbContext<DataContext>(options =>
                {
                    options.UseSqlServer(config.GetConnectionString("DefaultConnection"));
                });

            services.AddCors();
            services.AddScoped<IExportService, ExportService>();
            services.AddScoped<IImportService, ImportService>();
            services.AddScoped<IStyleService, StyleService>();
            services.AddScoped<IUpdateService, UpdateService>();
            services.AddAutoMapper(AppDomain.CurrentDomain.GetAssemblies());

            return services;
        }
    }
}