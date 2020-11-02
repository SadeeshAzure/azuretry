using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;

namespace FunctionApp2017
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {

            log.Info("C# HTTP trigger function processed a request.");
            string clientId = "";
            var res = new HttpResponseMessage(HttpStatusCode.OK);
            log.Info(JsonConvert.SerializeObject(req.Headers));
            try
            {
                clientId = req.Headers.First(y => y.Key.ToLower() == "x-adobesign-clientid").Value.First();

                //Json body to string variable
                StreamReader reader = new System.IO.StreamReader(await req.Content.ReadAsStreamAsync());
                reader.BaseStream.Position = 0;
                string requestFromPost = reader.ReadToEnd();

                //Extract file from body
                dynamic x = JsonConvert.DeserializeObject(requestFromPost);
                if (requestFromPost.Contains("agreement"))
                {
                    string agreementId = x.agreement.id.ToString();
                    string name = x.agreement.name.ToString();
                    var k = x.agreement.signedDocumentInfo.document.ToString();

                    byte[] docByteArr = Convert.FromBase64String(k);


                    byte[] auditByteArr = GetAuditReport(agreementId);

                    StoreinSharepoint(agreementId, name, docByteArr, auditByteArr);
                    //StoreInBlob(docByteArr); //Saving it in Azure Blob storage. Should be replaced with sharepoint store logic


                }
                if (String.IsNullOrEmpty(clientId))
                {
                    clientId = "ClientIDMissing";
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }


            log.Info("ClientID: " + clientId);

            res.Headers.Add("x-adobesign-clientid", clientId); //Mandatory to add the clientid in the response header. Else Adobe webhook wont work.

            return res;
        }

        private static byte[] GetAuditReport(string agreementId)
        {
            string auditAdobeSignURL = "https://api.eu1.documents.adobe.com:443/api/rest/v6/agreements/" +
                                agreementId + "/auditTrail";
            using (WebClient client = new WebClient())
            {
                client.Headers.Add("Authorization", "Bearer 3AAABLblqZhD_FuSyM0mhqeguqTMNGSwwqBmz2WlKWvg4oioNnY84XrSPTnBoO13KYMF0AtgzQpiAvVTEhgWszUUcpos4-Wqc");
                return client.DownloadData(auditAdobeSignURL);
            }
        }

        private async static void StoreInBlob(byte[] docByteArr)
        {
            string connectionString = "DefaultEndpointsProtocol=https;AccountName=blobstorageadobe;AccountKey=mz2GIdXeWCGzh0xiMsBeVI2/+/uB6iTvVE0DI6ebtex1NiNhAK4J6h3d/+9ZU4PG1DKvrevsOBDyzjiZLX1iFQ==;EndpointSuffix=core.windows.net";
            CloudStorageAccount storageAccount;
            storageAccount = CloudStorageAccount.Parse(connectionString);

            CloudBlobClient client;
            CloudBlobContainer container;

            client = storageAccount.CreateCloudBlobClient();

            container = client.GetContainerReference("signed-document");

            await container.CreateIfNotExistsAsync();
            CloudBlockBlob blob;

            string uniqueIdent;

            uniqueIdent = Guid.NewGuid().ToString("n");

            blob = container.GetBlockBlobReference(uniqueIdent + "_SignedDoc.pdf");
            blob.Properties.ContentType = "application/pdf";

            await blob.UploadFromByteArrayAsync(docByteArr, 0, docByteArr.Length);
        }


        private static async void StoreinSharepoint(string agreementId, string name, byte[] docByteArr, byte[] auditByteArr)
        {

            try
            {
                // Starting with ClientContext, the constructor requires a URL to the server running SharePoint.   
                using (ClientContext client = new ClientContext("https://tmhe.sharepoint.com/sites/TMHFR-AdobeDemo/"))
                {
                    client.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    var securepassword = new SecureString();
                    foreach (char c in "uc?4%wGp") { securepassword.AppendChar(c); }
                    client.Credentials = new SharePointOnlineCredentials("sadeeshkumar.ravanan.ext@toyota-industries.eu", securepassword);

                    // Assume that the web site has a library named "FormLibrary".   
                    var formLib = client.Web.Lists.GetByTitle("Finance Documents");
                    client.Load(formLib.RootFolder);
                    client.ExecuteQuery();

                    // FormTemplate path, The path should be on the local machine/server !  

                    Stream stream = new MemoryStream(docByteArr);
                    string uniqueIdent = Guid.NewGuid().ToString("n");


                    var fileUrl = String.Format("{0}/SignedDoc_{1}_{2}.pdf", formLib.RootFolder.ServerRelativeUrl, name, uniqueIdent);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(client, fileUrl, stream, true);
                    client.ExecuteQuery();

                    Stream stream2 = new MemoryStream(auditByteArr);
                    var fileUrl2 = String.Format("{0}/AuditReport_{1}_{2}.pdf", formLib.RootFolder.ServerRelativeUrl, name, uniqueIdent);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(client, fileUrl2, stream2, true);
                    client.ExecuteQuery();

                    
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }


}
