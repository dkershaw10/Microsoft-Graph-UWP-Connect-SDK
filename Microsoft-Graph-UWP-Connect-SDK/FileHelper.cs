//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Net.Http.Headers;

namespace Microsoft_Graph_UWP_Connect_SDK
{
    class FileHelper
    {
        /// <summary>
        /// Get a set of files.
        /// </summary>
        /// <param name="subject">The subject line of the email.</param>
        /// <returns>List of file</returns>
        internal async Task<DriveItem> GetFilesAsync()
        {
            try
            {
                // Initialize a new Graph client using a DelegateAuthenticationProvider to get an access token
                // Make sure to pass the current signed in user into the token getter, to ensure the right token is pulled from the cache
                GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
                {
                    AuthenticationResult ar = await App.PCApplication.AcquireTokenSilentAsync(App.initialScope,App.currentUser) ;
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ar.Token);
                }));

                // Get the driveItem representing expanded children of the root drive, and return
                DriveItem fileList = await graphClient.Me.Drive.Root.Request().Expand("children($select=id,name,file,webUrl)").GetAsync();
                return fileList;
            }


            catch (Exception e)
            {
                throw new Exception("We could not get files: " + e.Message);
            }
        }

    }
}
