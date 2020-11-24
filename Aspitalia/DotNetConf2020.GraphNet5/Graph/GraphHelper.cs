using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DotNetConf2020.GraphNet5
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }
        public static async Task<User> GetMeAsync()
        {
            try
            {
                return await graphClient.Me.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }
        public static async Task SendMailAsync()
        {
            try
            {
                var message = new Message
                {
                    Subject = "Testing from DotNetConf 2020",
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Text,
                        Content = "The Graph SDK is great!"
                    },
                    ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = "eba@expertsinside.com"
                            }
                        }
                    }
                };

                await graphClient.Me
                    .SendMail(message, true)
                    .Request()
                    .PostAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
            }
        }

        public static async Task<IUserJoinedTeamsCollectionPage> GetTeamsAsync()
        {
            try
            {
                var teams = await graphClient.Me.JoinedTeams
                    .Request()
                    .GetAsync();

                return teams;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }
    }
}