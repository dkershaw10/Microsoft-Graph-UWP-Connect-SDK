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
    class MailHelper
    {

        /// <summary>
        /// Get the signed-in users contact list
        /// </summary>
        /// <returns>First page of the users contact list</returns>
        internal async Task<IList<Contact>> GetContactsAsync()
        {
            try
            {
                // Initialize a new Graph client using a DelegateAuthenticationProvider to get an access token
                // Make sure to pass the current signed in user into the token getter, to ensure the right token is pulled from the cache
                GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
                {
                    AuthenticationResult ar = await App.PCApplication.AcquireTokenSilentAsync(App.initialScope, App.currentUser);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ar.Token);
                }));

                // get the contacts paged collection and return the CurrentPage (for demo only).  Normally this would be a search/picker experience.
                /// ADD CODE HERE TO GET AND RETURN CONTACTS
                var contacts = await graphClient.Me.Contacts.Request().Select("displayName, emailAddresses").GetAsync();
                return contacts.CurrentPage;
            }


            catch (Exception e)
            {
                throw new Exception("We could not get the user's contacts: " + e.Message);
            }
        }

        /// <summary>
        /// Compose and send a new email.
        /// </summary>
        /// <param name="subject">The subject line of the email.</param>
        /// <param name="bodyContent">The body of the email.</param>
        /// <param name="recipient">A single contact.</param>
        /// <returns></returns>
        internal async Task ComposeAndSendMailAsync(string subject,
                                                            string bodyContent,
                                                            string recipient)
        {

            // Prepare the recipient list and Add a new Recipient EmailAddress
            List<Recipient> recipientList = new List<Recipient>();
            recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = recipient.Trim() } });

            try
            {
                // Initialize a new Graph client using a DelegateAuthenticationProvider to get an access token
                // Make sure to pass the current signed in user into the token getter, to ensure the right token is pulled from the cache
                GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
                {
                    AuthenticationResult ar = await App.PCApplication.AcquireTokenSilentAsync(App.initialScope, App.currentUser);
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ar.Token);
                }));

                // create the message (with Subject, ToRecipients and Body/ItemBody
                var email = new Message
                {
                    Body = new ItemBody
                    {
                        Content = bodyContent,
                        ContentType = BodyType.Html
                    },
                    Subject = subject,
                    ToRecipients = recipientList,
                };

                // send the message
                try
                {
                    await graphClient.Me.SendMail(email, true).Request().PostAsync();
                }
                catch (ServiceException exception)
                {
                    throw new Exception("We could not send the message: " + exception.Error == null ? "No error message returned." : exception.Error.Message);
                }


            }

            catch (Exception e)
            {
                throw new Exception("We could not send the message: " + e.Message);
            }
        }
    }
}
