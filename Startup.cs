using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using TestAPI.Models;
using Newtonsoft.Json;

namespace TestAPI
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            //依赖注入
            services.AddDbContext<ProductContext>(options =>
            {
                options.UseSqlServer(Configuration.GetConnectionString("product"));
            });


            //添加cors 服务 配置跨域处理   
            services.AddCors(options =>
            {
                options.AddPolicy("any", builder =>
                {
                    builder.AllowAnyMethod().AllowAnyHeader()
                    //.AllowCredentials()//指定处理cookie
                    .AllowAnyOrigin();
                    //允许任何来源的主机访问
                });
            });

            services.AddControllers();

            services.AddControllers().AddNewtonsoftJson();
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseHttpsRedirection();

            app.UseRouting();

            app.UseAuthorization();

            // 注意，这一行代码一定要写在 app.UseAuthorization(); 的下面，否则会报错的。
            app.UseCors("any");
            

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();
            });
        }
    }
}
