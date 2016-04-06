//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Net.Http.Headers;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Windows.Storage;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Collections.ObjectModel;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Microsoft_Graph_UWP_Connect_SDK
{

    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private string _displayName = null;
        private MailHelper _mailHelper = new MailHelper();
        private FileHelper _fileHelper = new FileHelper();
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;
        private ObservableCollection<App.FileItem> fileItems = new ObservableCollection<App.FileItem>();
        private ObservableCollection<App.ContactItem> contactItems = new ObservableCollection<App.ContactItem>();

        public MainPage()
        {
            this.InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (!App.Current.Resources.ContainsKey("ida:ClientID"))
            {
                InfoText.Text = ResourceLoader.GetForCurrentView().GetString("NoClientIdMessage");
                ConnectButton.IsEnabled = false;
            }
            else
            {
                InfoText.Text = ResourceLoader.GetForCurrentView().GetString("ConnectPrompt");
                ConnectButton.IsEnabled = true;
            }
        }

        /// <summary>
        /// Signs in the current user
        /// </summary>
        /// <returns></returns>
        public async Task<bool> SignInCurrentUserAsync()
        {
            // Initialize a new Graph client using a DelegateAuthenticationProvider to get an access token
            // Also set the current signed in user
            GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(async (request) =>
            {
                AuthenticationResult ar = await App.PCApplication.AcquireTokenAsync(App.initialScope);
                App.currentUser = ar.User;
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ar.Token);
            }));

            var request = graphClient.Me.MailFolders.Inbox.Messages.Request().Top(pageSize.Value);
            // Get /me and save the signed-in user's _displayName
            /// ADD CODE HERE TO MAKE FIRST CALL TO GRAPH
            if (graphClient != null)
            {
                var user = await graphClient.Me.Request().GetAsync();
                _displayName = user.DisplayName;
                return true;
            }
            else
            {
                return false;
            }
                
        }

        
        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar.Visibility = Visibility.Visible;
            if (await SignInCurrentUserAsync())
            {
                InfoText.Text = "Hi " + _displayName + "," + Environment.NewLine + ResourceLoader.GetForCurrentView().GetString("SendMailPrompt");
                MailButton.IsEnabled = true;
                AttachText.Visibility = Visibility.Visible;
                ContactText.Visibility = Visibility.Visible;
                ConnectButton.Visibility = Visibility.Collapsed;
                DisconnectButton.Visibility = Visibility.Visible;
                MailButton.Visibility = Visibility.Collapsed;
                MailButton.Visibility = Visibility.Visible;
                FileCombo.Visibility = Visibility.Visible;
                ContactCombo.Visibility = Visibility.Visible;

                // Load the oneDrive files
                await GetFiles();

                // Load the signed-in user's contacts
                await GetContacts();
            }
            else
            {
                InfoText.Text = ResourceLoader.GetForCurrentView().GetString("AuthenticationErrorMessage");
            }

            ProgressBar.Visibility = Visibility.Collapsed;
        }

        // on click call the mailhelper module to send mail
        private async void MailButton_Click(object sender, RoutedEventArgs e)
        {
            App.FileItem selectedFile = fileItems[FileCombo.SelectedIndex];
            App.ContactItem selectedContact = contactItems[ContactCombo.SelectedIndex];
            ProgressBar.Visibility = Visibility.Visible;
            MailStatus.Text = string.Empty;
            try
            {
                await _mailHelper.ComposeAndSendMailAsync(ResourceLoader.GetForCurrentView().GetString("MailSubject"), ComposePersonalizedMail(selectedContact.name, selectedFile.webUrl), selectedContact.address);
                MailStatus.Visibility = Visibility.Visible;
                MailStatus.Text = string.Format(ResourceLoader.GetForCurrentView().GetString("SendMailSuccess"), selectedContact.name);
            }
            catch (Exception)
            {
                MailStatus.Visibility = Visibility.Visible;
                MailStatus.Text = ResourceLoader.GetForCurrentView().GetString("MailErrorMessage");
            }
            finally
            {
                ProgressBar.Visibility = Visibility.Collapsed;
            }
            
        }

        // <summary>
        // Personalizes the email.
        // </summary>
        public static string ComposePersonalizedMail(string userName, string link)
        {
            return String.Format(ResourceLoader.GetForCurrentView().GetString("MailContents"), userName, link);
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar.Visibility = Visibility.Visible;
            //workaround
            App.PCApplication.UserTokenCache.Clear(App.PCApplication.ClientId);
            // not working yet
            //foreach (var user in App.PCApplication.Users)
            //{
            //    user.SignOut();
            //}
            ProgressBar.Visibility = Visibility.Collapsed;
            MailButton.IsEnabled = false;
            MailButton.Visibility = Visibility.Collapsed;
            FileCombo.Visibility = Visibility.Collapsed;
            ContactCombo.Visibility = Visibility.Collapsed;
            AttachText.Visibility = Visibility.Collapsed;
            ContactText.Visibility = Visibility.Collapsed;
            MailStatus.Visibility = Visibility.Collapsed;
            ConnectButton.Visibility = Visibility.Visible;
            InfoText.Text = ResourceLoader.GetForCurrentView().GetString("ConnectPrompt");
            this._displayName = null;
            fileItems.Clear();
            contactItems.Clear();
        }

        // call the filehelper module to get the user's files
        private async Task GetFiles()
        {
            var files = await _fileHelper.GetFilesAsync();
            foreach (DriveItem d in files.Children.CurrentPage)
            {
                if (d.File != null)
                {
                    var fileItem = new App.FileItem
                    {
                        name = d.Name,
                        webUrl = d.WebUrl
                    };
                    fileItems.Add(fileItem);
                }
            }
        }

        // call the mailhelper module to get the user's contacts
        private async Task GetContacts()
        {
            var contacts = await _mailHelper.GetContactsAsync();
            foreach (Contact c in contacts)
            {
                var emailAddress = "";
                foreach (var _address in c.EmailAddresses)
                {
                    emailAddress = _address.Address;
                }
                var contactItem = new App.ContactItem
                {
                    name = c.DisplayName,
                    address = emailAddress
                };
                contactItems.Add(contactItem);
            }
        }
    }
}
