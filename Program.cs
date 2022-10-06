using Azure.Identity;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Graph
{
    public class Program
    {
        private static AppSettings ApplicationSettings = new();
        public static void Main()
        {
            Console.WriteLine("Hello, World!");
            var host = Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                services.AddSingleton(options =>
                {
                    var configuration = context.Configuration;

                    configuration.Bind(ApplicationSettings);
                    return configuration;
                });
                services.AddSingleton(sp =>
                {
                    return ApplicationSettings;
                });
                services.AddScoped(sp =>
                {
                    var options = new TokenCredentialOptions
                    {
                        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                    };
                    var certificate = LoadCertificate(ApplicationSettings);
                    var credentials = new ClientCertificateCredential(ApplicationSettings.TenantId, ApplicationSettings.ClientId, certificate, options);
                    return new GraphServiceClient(credentials);
                });
                services.AddSingleton<GraphServiceProcessor, GraphServiceProcessor>();
            })
            .Build();
            var processor = host.Services.GetRequiredService<GraphServiceProcessor>();
            var result = processor.GetListItems().GetAwaiter().GetResult();
            System.Console.WriteLine(result.Count);
            processor.AddTestListItem().GetAwaiter().GetResult();
            var drives = processor.GetDrives().GetAwaiter().GetResult();
            System.Console.WriteLine(drives.Count);
        }

        private static X509Certificate2 LoadCertificate(AppSettings appSettings)
        {
            // Will only be populated correctly when running in the Azure Function host
            string certBase64Encoded = appSettings.CertificateFromKeyVault;

            if (!string.IsNullOrEmpty(certBase64Encoded))
            {
                // Azure Function flow
                return new X509Certificate2(Convert.FromBase64String(certBase64Encoded),
                                            "",
                                            X509KeyStorageFlags.Exportable |
                                            X509KeyStorageFlags.MachineKeySet |
                                            X509KeyStorageFlags.EphemeralKeySet);
            }
            else
            {
                // Local flow
                var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                var certificateCollection = store.Certificates.Find(X509FindType.FindByThumbprint, appSettings.CertificateThumbprint, false);
                store.Close();

                return certificateCollection.First();
            }
        }
    }
}
