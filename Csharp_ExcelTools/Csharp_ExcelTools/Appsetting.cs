public static class Appsetting
{
    public static IConfiguration GetConfigurations()
    {
        IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile(@"Configuraction\appsettings.json")
            .AddEnvironmentVariables()
            .Build();

        return config;
    }

}