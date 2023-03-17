namespace HPAutomation
{
    using HPAutomation.Application;
    using HPAutomation.Logging;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;

    public static class Program
    {
        // AccessDatabaseEngine_X64 -- Might be needed if you get any exception
        public static void Main(string[] args)
        {

            var host = Host.CreateDefaultBuilder(args)
                          .ConfigureServices((_, services) =>
                              services.AddScoped<DataMigrationService>())
                          .Build();

            using (var serviceScope = host.Services.CreateScope())
            {
                try
                {
                    var services = serviceScope.ServiceProvider;
                    var dataMigrationService = services.GetRequiredService<DataMigrationService>();
                    LogFile.SetFilePath();

                    LogFile.WriteLog("Log: Migration Started");

                    dataMigrationService.MigrateData();

                    LogFile.WriteLog("Log: Migration Completed");
                }
                catch (Exception ex)
                {
                    LogFile.WriteLog("Error: An error occurred while migrating the data" + ex.Message);
                }

                Console.ReadLine();
            }
        }
    }
}