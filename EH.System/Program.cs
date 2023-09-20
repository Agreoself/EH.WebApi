using EH.System.Commons; 
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Implement;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using EH.Service.Implement;
using EH.Service.Interface;
using Microsoft.AspNetCore.Authentication.Negotiate;
using Microsoft.EntityFrameworkCore;
using NLog.Web;
using EH.System.DIExtension;
using Microsoft.Extensions.Options;
using static EH.Repository.DataAccess.MyAppDbContext;

namespace EH.System
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            var Configuration = builder.Configuration;
            var services = builder.Services;
         
            services.AddDbContext<MyAppDbContext>(opt =>
            {
                var connecStr = Configuration.GetValue<string>("ConnectionString:DefaultConnection");
                opt.UseSqlServer(connecStr);
                opt.AddInterceptors(new CreateByInterceptor(services.BuildServiceProvider().GetRequiredService<IHttpContextAccessor>()));
            });

            services.AddLogging(logger =>
            {
                logger.ClearProviders();//删除所有其他的关于日志记录的配置
                logger.SetMinimumLevel(LogLevel.Information);//设置最低的log级别
                logger.AddNLog(Configuration.GetValue<string>("NLog:Path"));//支持nlog 
            });

            services.AddCors(policy =>
            {
                policy.AddPolicy("CorsPolicy", opt => opt
                .WithOrigins("http://localhost:8080")
                .AllowCredentials()
                .AllowAnyHeader()
                .AllowAnyMethod());
            });

            services.AddControllers();

            services.AddAuthentication(NegotiateDefaults.AuthenticationScheme).AddNegotiate();

            services.AddAuthorization(options =>
            {
                // By default, all incoming requests will be authorized according to the default policy.
                options.FallbackPolicy = options.DefaultPolicy;
            });

            services.AddSingleton(typeof(JwtHelper));
            services.AddSingleton(new LogHelper());

            services.RegisterAllServices();

            //services.AddTransient<ISysMenusRepository, SysMenusRepository>();
            //services.AddTransient<ISysMenuService, SysMenuService>();
            //services.AddTransient<ISysRolesRepository, SysRolesRepository>();
            ////services.AddTransient<ISysUserService, SysUserService>();
            //services.AddTransient<ISysUserRoleRepository, SysUserRoleRepository>();
            ////services.AddTransient<ISysUserService, SysUserService>();
            //services.AddTransient<ISysRoleMenuRepository, SysRoleMenuRepository>();
            ////services.AddTransient<ISysUserService, SysUserService>();
            //services.AddTransient<ISysUsersRepository,SysUsersRepository>();
            //services.AddTransient<ISysUserService, SysUserService>();
            //services.AddTransient<IADUserService, ADUserService>();
     
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            services.AddEndpointsApiExplorer();
            services.AddSwaggerGen();

            var app = builder.Build();
            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseCors("CorsPolicy");

            app.UseAuthentication();
            app.UseAuthorization();
       
            app.MapControllers();
            app.Run();
        }
    }
}