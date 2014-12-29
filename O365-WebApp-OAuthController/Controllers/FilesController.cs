//----------------------------------------------------------------------------------------------
//    Copyright 2014 Microsoft Corporation
//
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using O365_WebApp_OAuthController.Models;
using O365_WebApp_OAuthController.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace O365_WebApp_OAuthController.Controllers
{

    [Authorize]
    public class FilesController : Controller
    {
        // GET: Files
        public async Task<ActionResult> Index(string authError)
        {
            List<MyFile> myFiles = new List<MyFile>();

            AuthenticationContext authContext = null;
            AuthenticationResult result = null;
            CapabilityDiscoveryResult filesCapabilityDiscoveryResult = null;
            DiscoveryClient discoveryClient = null;
            SharePointClient spClient = null;

            string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/nameidentifier").Value;

            ClientCredential credential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.AppKey);

            try
            {
                authContext = new AuthenticationContext(SettingsHelper.Authority, new TokenDbCache(userObjectID));

                if (authError != null)
                {
                    Uri redirectUri = new Uri(Request.Url.GetLeftPart(UriPartial.Authority).ToString() + "/OAuth");

                    string state = GenerateState(userObjectID, Request.Url.ToString());

                    ViewBag.AuthorizationUrl = authContext.GetAuthorizationRequestURL(SettingsHelper.DiscoveryServiceResourceId,
                        SettingsHelper.ClientId,
                        redirectUri,
                        UserIdentifier.AnyUser,
                        state == null ? null : "&state=" + state);

                    ViewBag.ErrorMessage = "UnexpectedError";

                    return View(myFiles);
                }

                // Query Discovery Service to retrieve MyFiles capability result
                result = authContext.AcquireTokenSilent(SettingsHelper.DiscoveryServiceResourceId, credential, UserIdentifier.AnyUser);
                discoveryClient = new DiscoveryClient(() =>
                {
                    return result.AccessToken;
                });
                
            }
            catch (AdalException e)
            {
                if (e.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {

                    // The user needs to re-authorize.  Show them a message to that effect.
                    // If the user still has a valid session with Azure AD, they will not be prompted for their credentials.                 
                    ViewBag.ErrorMessage = "AuthorizationRequired";
                    authContext = new AuthenticationContext(SettingsHelper.Authority);
                    Uri redirectUri = new Uri(Request.Url.GetLeftPart(UriPartial.Authority).ToString() + "/OAuth");

                    string state = GenerateState(userObjectID, Request.Url.ToString());

                    ViewBag.AuthorizationUrl = authContext.GetAuthorizationRequestURL(SettingsHelper.DiscoveryServiceResourceId,
                        SettingsHelper.ClientId,
                        redirectUri,
                        UserIdentifier.AnyUser,
                        state == null ? null : "&state=" + state);

                    return View(myFiles);
                }

                ViewBag.ErrorMessage = "Error while Acquiring Token from Cache.";

                return View("Error");

            }

            try
            {
                ViewBag.Office365User = result.UserInfo.GivenName;

                ActiveDirectoryClient adGraphClient = new ActiveDirectoryClient(SettingsHelper.AADGraphServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.AADGraphResourceId, credential, UserIdentifier.AnyUser);

                        return authResult.AccessToken;
                    });

                var currentUser = await adGraphClient.Users.Where(u => u.ObjectId == result.UserInfo.UniqueId).ExecuteSingleAsync();

                ViewBag.Office365User = String.Format("{0} ({1})",currentUser.DisplayName,currentUser.Mail);

                filesCapabilityDiscoveryResult = await discoveryClient.DiscoverCapabilityAsync("MyFiles");

                // Initialize SharePoint client to query users' files
                spClient = new SharePointClient(filesCapabilityDiscoveryResult.ServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(filesCapabilityDiscoveryResult.ServiceResourceId, credential, UserIdentifier.AnyUser);

                        return authResult.AccessToken;
                    });
               
                // Query users' files and get the first paged collection
                var filesCollection = await spClient.Files.ExecuteAsync();
                var files = filesCollection.CurrentPage;
                foreach (var file in files)
                {
                    myFiles.Add(new MyFile { Name = file.Name });
                }

                return View(myFiles);
            }
            catch (Exception e)
            {
                ViewBag.ErrorMessage = String.Format("UnexpectedError: {0}", e.Message);

                return View("Error");
            }           
        }

        /// Generate a state value using a random Guid value and the origin of the request.
        /// The state value will be consumed by the OAuth controller for validation and redirection after login.
        /// Here we store the random Guid in the database cache for validation by the OAuth controller.
        public string GenerateState(string userObjId, string requestUrl)
        {
            try
            {
                string stateGuid = Guid.NewGuid().ToString();
                ApplicationDbContext db = new ApplicationDbContext();
                db.UserStateValues.Add(new UserStateValue { stateGuid = stateGuid, userObjId = userObjId });
                db.SaveChanges();

                List<String> stateList = new List<String>();
                stateList.Add(stateGuid);
                stateList.Add(requestUrl);

                var formatter = new BinaryFormatter();
                var stream = new MemoryStream();
                formatter.Serialize(stream, stateList);
                var stateBits = stream.ToArray();

                return Url.Encode(Convert.ToBase64String(stateBits));
            }
            catch
            {
                return null;
            }

        }
    }
}