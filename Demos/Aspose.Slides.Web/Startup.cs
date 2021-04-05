using System.IO;
using Aspose.Slides.Web.API.Json;
using Aspose.Slides.Web.Core.DependencyInjection;
using Aspose.Slides.Web.UI.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Aspose.Slides.Web
{
    public class Startup
    {
        public Startup(IConfiguration configuration, IWebHostEnvironment env)
        {
            Configuration = configuration;
            Env = env;
        }

        public IWebHostEnvironment Env { get; set; }

        public IConfiguration Configuration { get; set; }


        public void ConfigureServices(IServiceCollection services)
        {
            services
                .AddControllers()
                .AddJsonOptions(
                    options =>
                    {
                        // PascalCase serialization is turned on
                        options.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
                        options.JsonSerializerOptions.PropertyNamingPolicy = null;
                        options.JsonSerializerOptions.IgnoreNullValues = true;
                        options.JsonSerializerOptions.Converters.Add(new BooleanNullHandlingConverter());
                    }
                );

            services.AddControllersWithViews();

            services.AddScoped<ISlidesViewModelFactory, SlidesViewModelFactory>();
            services.AddScoped<IEditorService, EditorService>();

            var rootDir = Path.Combine(Path.GetDirectoryName(typeof(Startup).Assembly.Location), "temp");
            services.AddLocalFileStorage(Path.Combine(rootDir, "storages"));
            services.AddSlidesServices(
                Path.Combine(rootDir, "tmp"),
                "odp,otp,pptx,pptm,potx,ppt,pps,ppsm,pot,potm,pdf,xps,ppsx,tiff,html,swf,doc,docx,bmp,jpeg,jpg,png,emf,wmf,gif,exif,ico,svg,xls,xlsx",
                "<path to ffmpeg binaries>",
                "<path to Aspose license>");
            services.AddPresentationTemplates("http://example.com");
        }

        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            app.UseStaticFiles();
            app.UseRouting();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllers();

                endpoints.MapControllerRoute(
                    "SlidesNavigationRoute",
                    "slides/{toaction}/navigation/{*values}",
                    new { controller = "Slides", action = "PermanentlyRedirect" }
                );
                endpoints.MapControllerRoute(
                    "SlidesNavigationExtRoute",
                    "slides/{toaction}/{extension}/navigation/{*values}",
                    new { controller = "Slides", action = "PermanentlyRedirect" }
                );
                endpoints.MapControllerRoute(
                    "SlidesRoute",
                    "slides/{action}/{extension}",
                    new { controller = "Slides" }
                );

                endpoints.MapControllerRoute(
                    "default",
                    "{controller=Home}/{action=Index}"
                );
            });
        }
    }
}
