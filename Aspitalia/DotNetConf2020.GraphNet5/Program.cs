using Microsoft.Extensions.Configuration;
using System;

namespace DotNetConf2020.GraphNet5
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            var appConfig = LoadAppSettings();
            
            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');
            var authProvider = new DeviceCodeAuthProvider(appId, scopes);
            
            var accessToken = await authProvider.GetAccessToken();
            Console.WriteLine($"Access token: {accessToken}\n");
            GraphHelper.Initialize(authProvider);

            var user = await GraphHelper.GetMeAsync();
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

            var teams = await GraphHelper.GetTeamsAsync();

            foreach (var team in teams)
            {
                Console.WriteLine($"Team: {team.DisplayName}!\n");
            }

            await GraphHelper.SendMailAsync();

            Console.WriteLine("Email sent!");

            Console.ReadLine();
        }

        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }
            return appConfig;
        }
    }
}
