namespace HPAutomation.Infrastructure
{
    using Microsoft.Extensions.Configuration;
    using System.IO;

    public class AppSettings
    {
        static IConfiguration configuration;

        static AppSettings()
        {
            configuration = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).Build();

        }

        public static string DbConnectionString
        {
            get
            {
                return configuration.GetSection("ConnectionStrings")["DBConnectionString"];
            }
        }

        public static string FileLocation
        {
            get
            {
                return configuration.GetSection("FileSettings")["Location"];
            }
        }
    }
}
