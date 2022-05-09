using Azure.Identity;

using Microsoft.Graph;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace MSGraph_SDK_Net
{
    internal class Program
    {
        /// <summary>
        /// This project using the client credentials grant flow for the authentication provider for Microsoft Graph SDK.
        /// You must configure an app registration in your tenant with the Microsoft Graph Application permissions of User.ReadWrite.All
        /// to run this project as is.  Please populate the tenantId, clientId, clientSecret, and your own extensionAttributes
        /// </summary>
        static string tenantId = "";
        static string clientId = "";       
        static string clientSecret { 
            get { 
                return ""; 
            } 
        }

        static class ExtensionAttributes
        {
            public static string primaryContact = "";
        }


        static ClientSecretCredential _authProvider = null;
        static ClientSecretCredential AuthProvider
        {
            get
            {
                if(_authProvider == null)
                {
                    _authProvider = new ClientSecretCredential(tenantId, clientId, clientSecret);
                }
                return _authProvider;
            }
        }

        static GraphServiceClient _graphServiceClient = null;
        static GraphServiceClient GraphServiceClient
        {
            get
            {
                if(_graphServiceClient == null)
                {
                    _graphServiceClient = new GraphServiceClient(AuthProvider, new List<String>() { "https://graph.microsoft.com/.default" });
                }
                return _graphServiceClient;
            }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Enter the UPN of the user you want to modify:");
            string upn = Console.ReadLine();

            Task<string> t = GetExtensionAttributeValue(upn, ExtensionAttributes.primaryContact);
            Console.WriteLine($"Value for '{ExtensionAttributes.primaryContact}': '{t.Result}'");

            Console.WriteLine($"\nEnter the new value for '{ExtensionAttributes.primaryContact}'");
            string newValue = Console.ReadLine();

            Task<List<string>> s = UpdateOrAddExtensionAttributeValue(upn, ExtensionAttributes.primaryContact, newValue);
            foreach(string value in s.Result)
            {
                Console.WriteLine(value);
            }

            Console.WriteLine($"\nPress any key to quit...");
            Console.ReadKey();

        }


        static async Task<List<string>> UpdateOrAddExtensionAttributeValue(string upn, string extensionName, string value)
        {
            List<string> status = new List<string>();

            User u = await GraphServiceClient.Users[upn]
                .Request()
                .Select($"id,displayName,{extensionName}")
                .GetAsync();

            if (!u.AdditionalData.ContainsKey(extensionName))
            {
                // the extension wasn't on the user so we are going to add it to the AdditionalData collection before saving
                u.AdditionalData.Add(extensionName, value);
                status.Add($"Added the value for '{u.AdditionalData[extensionName]}' to the user '{upn}'.");
            } else
            {
                // set the new value on the user object in the Additional Data Field
                u.AdditionalData[extensionName] = value; 
                status.Add($"Changed the value for '{extensionName}' to '{u.AdditionalData[extensionName]}' for user '{upn}'");
            }

            // save the change on the user object
            try
            {
                await GraphServiceClient.Users[u.Id].Request().UpdateAsync(u);
                status.Add($"\nSuccessfully updated the user '{upn}'");
            } catch (ServiceException e)
            {
                status.Add($"MS Graph error: {e.Error}");
            }
            return status;
        }


        static async Task<string> GetExtensionAttributeValue(string upn, string extensionName)
        {
            User u = await GraphServiceClient.Users[upn]
                .Request()
                .Select($"id,{extensionName}")
                .GetAsync();

            try
            {
                // the additional data field on the user object has the extension attributes
                return u.AdditionalData[extensionName].ToString(); 
            } catch
            {
                // if the property isn't found, a null exception key not found type of error occurs, just return an empty string
                return string.Empty;
            }
        }
    }
}
