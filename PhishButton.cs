using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

// The phish alert button code.
namespace SSQPhishAlert
{
    [ComVisible(true)]

    // Setup the button
    public class PhishButton : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private string reporter;

        public PhishButton()
        {
        }

        #region IRibbonExtensibility Members

        // Phish button XML. This is the XML file read by Office that creates the button
        // in oulook. Refer to the PhishButton.xml for the settings there
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SSQPhishAlert.PhishButton.xml");
        }

        // Code for when the button is pressed. 
        public void OnTextButton(Office.IRibbonControl control)
        {

            // Get the junk folder and map it to the junk object
            Outlook.MAPIFolder Junk = (Outlook.MAPIFolder)Globals.PhishAlert.Application.ActiveExplorer().
                Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);

            // catch all string message should none of the conditions meet. 
            String itemMessage = "Email not selected!";

            // try and execute the code. should it fail, then popup with the catch all message
            try
            {
                if (Globals.PhishAlert.Application.ActiveExplorer().Selection.Count == 1 )
                {
                    
                    
                    bool ValidFolder = true;

                    // get current folder of the selected email
                    Object CurrentFolder = Globals.PhishAlert.Application.ActiveExplorer().CurrentFolder.Name;

                    // get the first selected item. 
                    Object selObject = Globals.PhishAlert.Application.ActiveExplorer().Selection[1];

                    // Check to make sure its not selected from the following folders. We dont care about
                    // These folders as they are either already flagged as spam, or not going to be a phish
                    if (CurrentFolder.ToString() == "Junk Email"
                        || CurrentFolder.ToString() == "Drafts"
                        || CurrentFolder.ToString() == "Sent Items"
                        || CurrentFolder.ToString() == "Deleted Items"
                        || CurrentFolder.ToString() == "Conversation History")
                    {
                        ValidFolder = false;
                        // Set the message text saying that its a bad folder
                        itemMessage = $"Cannot Submit Emails In The {CurrentFolder.ToString()} Folder";                     

                    }

                    // So we have a valid email and in not in a special folder (as stated above).
                    if (selObject is Outlook.MailItem && ValidFolder == true)
                    {

                        // Get the phish reporters email address. Used in the submission. 
                        Outlook.AddressEntry addrEntry = Globals.PhishAlert.Application.Session.CurrentUser.AddressEntry;

                        if (addrEntry.Type == "EX")
                        {
                            Outlook.ExchangeUser currentUser =
                                Globals.PhishAlert.Application.Session.CurrentUser.
                                AddressEntry.GetExchangeUser();

                            reporter = currentUser.PrimarySmtpAddress;
                        }

                        // Get the phishing email object
                        Outlook.MailItem mailItem = (selObject as Outlook.MailItem);

                        // Create a new email to send to phishreport
                        Outlook.MailItem eMail = (Outlook.MailItem)
                            Globals.PhishAlert.Application.CreateItem(Outlook.OlItemType.olMailItem);

                        // Set the contents of the new email. 
                        // Subject is the same as the phsihing email
                        // Attach the phsihing email
                     

                        eMail.Subject = mailItem.Subject;
                        eMail.To = "phishrecipient@domain.tld";
                        eMail.Body = $@"Phishing Report

affected_contact:{reporter}
requestor:{reporter}
short_desc:PHISHREPORT - {mailItem.Subject}
";
                        eMail.Attachments.Add(mailItem);
                        mailItem.Move(Junk);

                        // Send the email
                        ((Outlook._MailItem)eMail).Send();

                        //eMail.Display(false);

                        // Set the message text saying that it was submitted.
                        itemMessage = "Phishing Report Submitted :)";
                    }
                }

                // One of the catch conditions. If the human selected more than one message
                // tell them to select only one. 
                if (Globals.PhishAlert.Application.ActiveExplorer().Selection.Count > 1)
                {

                    itemMessage = "Please Select Only One Item" ;

                }

            // Next catch... I dunno why this is here. May need to be removed.
            }
            catch (Exception ex)
            {
                itemMessage = ex.Message;
            }

            // Send the status message to the human
            MessageBox.Show(itemMessage);
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
