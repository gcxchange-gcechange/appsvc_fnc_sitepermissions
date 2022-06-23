using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;

namespace SitePermissions
{
    public static class StoreData
    {
        public static async Task<bool> StoreReports(ExecutionContext context, List<Report> reports, string containerName, ILogger log)
        {
            CreateContainerIfNotExists(context, containerName);

            var storageAccount = GetCloudStorageAccount(context);
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference(containerName);

            var now = DateTime.Now;
            string FileTitle = now.ToString("dd-MM-yyyy") + "-" + containerName + ".json";

            var blob = container.GetBlockBlobReference(FileTitle);
            blob.Properties.ContentType = "application/json";

            var json = JsonConvert.SerializeObject(reports.ToArray());

            using (var ms = new MemoryStream())
            {
                LoadStreamWithJson(ms, json);
                await blob.UploadFromStreamAsync(ms);
            }

            log.LogInformation($"Blob {FileTitle} has been uploaded to container {container.Name}");

            await blob.SetPropertiesAsync();

            return true;
        }

        private static async void CreateContainerIfNotExists(ExecutionContext executionContext, string ContainerName)
        {
            var storageAccount = GetCloudStorageAccount(executionContext);
            var blobClient = storageAccount.CreateCloudBlobClient();
            string[] containers = new string[] { ContainerName };

            foreach (var item in containers)
            {
                var blobContainer = blobClient.GetContainerReference(item);
                await blobContainer.CreateIfNotExistsAsync();
            }
        }

        private static CloudStorageAccount GetCloudStorageAccount(ExecutionContext executionContext)
        {
            var config = new ConfigurationBuilder()
                            .SetBasePath(Environment.CurrentDirectory/*executionContext.FunctionAppDirectory*/)
                            .AddJsonFile("local.settings.json", true, true)
                            .AddEnvironmentVariables().Build();
            var storageAccount = CloudStorageAccount.Parse(config["AzureWebJobsStorage"]);
            return storageAccount;
        }

        private static void LoadStreamWithJson(Stream ms, object obj)
        {
            var writer = new StreamWriter(ms);
            writer.Write(obj);
            writer.Flush();
            ms.Position = 0;
        }
    }
}
