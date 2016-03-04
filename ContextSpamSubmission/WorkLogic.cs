using Ionic.Zip;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContextSpamSubmission
{
    class WorkLogic
    {
        const string PR_MAIL_HEADER_TAG = @"http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        const string PR_ATTACH_DATA_BIN = @"http://schemas.microsoft.com/mapi/proptag/0x37010102";
        //variable declarations XXX
        //the registry hive containing our address keys.
        string sRegPath = "HKEY_CURRENT_USER\\SOFTWARE\\InverseSoftware\\SpamSubmission\\";
        //the key containing the ticket 'voicemail' address. Emailing this address should result in
        //a ticket being created, with a reference to the SPAM sample.
        string sRegTicketAddress = "ticketEmail";
        //the key containing the address we submit the SPAM sample to.
        string sRegSubmitAddress = "spamEmail";
        //key that holds the zip password
        string sRegEncryptionPassword = "encryptionPassword";
        //string to store the registry key holding the Debug value
        string sRegDebug = "debug";
        //key to hold the ticket voicemail address, once we get it.
        string sEmailTicketAddress = "";
        //key to hold the SPAM submission address, once we get it.
        string sSpamSubmitAddress = "";
        //string to store encryption password
        string sEncryptionPassword = "";
        //A string to hold the interesting items we want to report on in plaintext
        string sMetadata = "";
        //boolean value (stored in reg) dictating wether or not we should show debugging messages
        bool bDebug = false;
        //an Outlook Rules Array to store all the current outlook rules.
        Rules olRuleList = null;
        //Single Rule instance
        Rule olRule = null;
        //A string for the rule name we will use. Registry?
        string olRuleName = "SPAMAutoDeleteList";
        //Boolean value for existence of SPAM Rule in outlook
        bool bSpamRuleExists = false;

        
        
        public void submit()
        {
            //get our reference to the application for future use.
            Microsoft.Office.Interop.Outlook.Application outlookApp = Globals.ThisAddIn.Application;

            //Main logic, majority of program logic is below, in this method.
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            //this checks something, and is probably important. Copy Paste form the interwebs.
            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                //item = variable storing what was right clicked on.
                object item = explorer.Selection[1];
                //if the item selected is a mail item, we know the user has done it right, let's proceed.
                if (item is MailItem)
                {
                    //store badmail for future use
                    MailItem badMail = item as MailItem;
                    if (badMail != null)
                    {
                        //set the guid of this message
                        string sGuid = Guid.NewGuid().ToString();
                        string sHeaders = "";
                        PropertyAccessor oPA = badMail.PropertyAccessor as PropertyAccessor;
                        
                        try
                        {
                            sHeaders = (string)oPA.GetProperty(PR_MAIL_HEADER_TAG);
                        }
                        catch(System.Exception e) { Console.WriteLine(e); }

                        //This will pull out the headers and such, and whack them into variables.
                        sMetadata = "To: " +  badMail.To + "\r\n";
                        sMetadata += "From: " +  badMail.SenderName + ": " + badMail.SenderEmailAddress + "\r\n";
                        sMetadata += "Subject: " + badMail.Subject + "\r\n";
                        sMetadata += "CC: " + badMail.CC + "\n\r";
                        sMetadata += "Companies Associated With Email: " + badMail.Companies + "\r\n";
                        sMetadata += "Email Creation Time: " + badMail.CreationTime + "\r\n";
                        sMetadata += "Delivery Report Requested: " +badMail.OriginatorDeliveryReportRequested + "\r\n";
                        sMetadata += "Received Time: " + badMail.ReceivedTime + "\r\n";
                        sMetadata += "Sent On: " + badMail.SentOn.ToString() + "\r\n";
                        sMetadata += "Size (kb): " + ((badMail.Size)/1024).ToString() + "\r\n";
                        sMetadata += "Headers: \r\n" + sHeaders + "\r\n";
                        sMetadata += "Plaintext Body: \r\n" + badMail.Body + "\r\n";

                        //This will create a mail item, and send it to the designated mailbox of a ticketing system.
                        MailItem ticketMail = (MailItem) outlookApp.CreateItem(OlItemType.olMailItem);
                        ticketMail.To = sEmailTicketAddress;
                        ticketMail.Subject = sGuid;
                        ticketMail.Body = sMetadata;
                        ticketMail.Send();


                        //Save the badmail to disk, to then read back in in a compressed stream.
                        
                        //First, get temp path(checks the below in order):
                        //The path specified by the TMP environment variable.
                        //The path specified by the TEMP environment variable.
                        //The path specified by the USERPROFILE environment variable.
                        //The Windows directory.
                        string tempDir = Path.GetTempPath();
                        string badOnDisk = tempDir + sGuid + ".msg";
                        string badZipOnDisk = tempDir + sGuid + ".zip";

                        //Save it to disk
                        badMail.SaveAs(badOnDisk);

                        //Read in the email in the .zip format, with a password, write back to disk.
                        using (ZipFile zip = new ZipFile())
                        {
                            zip.Password = sEncryptionPassword;
                            //the "." specifies the directory structure inside the zip - . just means
                            //insert the attachment at the root, instead of nested in a replication of
                            //the systems temp dir
                            zip.AddFile(badOnDisk, ".");
                            zip.Save(badZipOnDisk);
                        }

                        //Better get rid of the raw BadMail as soon as we're done with it
                        File.Delete(badOnDisk);

                        //This will create a mail item, and send it to a sample collection mailbox, with the badSample attached.
                        MailItem spamMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
                        spamMail.To = sSpamSubmitAddress;
                        spamMail.Subject = sGuid;
                        spamMail.Body = sMetadata;
                        spamMail.Attachments.Add(badZipOnDisk, OlAttachmentType.olByValue, 1, "SPAM Sample " + sGuid);
                        //spamMail.Send();

                        //That's sent, let's delete the .zip on disk
                        File.Delete(badZipOnDisk);

                        //if the listed email address doesn't contain an @, it's not a legit threat address, disregard blocking.
                        if (badMail.SenderEmailAddress.Contains("@"))
                        {
                            DialogResult blockSender = MessageBox.Show("Do you want to automatically delete future emails from:\n" + badMail.SenderEmailAddress, "Block Sender?", MessageBoxButtons.YesNo);
                            if (blockSender == DialogResult.Yes)
                            {
                                blacklistSender(badMail.SenderEmailAddress);
                            }
                        }
                                               
                        //Finally, remove the dodgy email from outlook.
                        badMail.UnRead = false;
                        badMail.Save();
                        badMail.Delete();

                        //If we're debugging, let's show the success and contents.
                        if (bDebug)
                        {
                            MessageBox.Show("You've submitted a SPAM sample.\r\n" +
                                sMetadata, "Thanks", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    
                    //Not a mail item, need to decide how to handle this. Advise user they done goofed.
                    //Should never happen, theoretically.
                    else
                    {
                        MessageBox.Show("You've selected something that is not an email.\r\n" + 
                            "Please ensure you right click the email you want to submit, and try again.\r\n\r\n" +
                            "If the issue persists, please contact the service centre for assistance.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
            }
        }
       
        public bool initialize()
        {
            //bDebug string for testing
            string debugMessage= "";
            //let's grab our stuff from the registry
            try
            {
                bDebug = Registry.GetValue(sRegPath, sRegDebug, "false").ToString().ToLower().Equals("true");

                sEmailTicketAddress = Registry.GetValue(sRegPath, sRegTicketAddress, null).ToString();
                debugMessage += "Email Ticket Address: " + sEmailTicketAddress + "\n";

                sSpamSubmitAddress = Registry.GetValue(sRegPath, sRegSubmitAddress, null).ToString();
                debugMessage += "Spam Submit Address: " + sSpamSubmitAddress + "\n";

                sEncryptionPassword = Registry.GetValue(sRegPath, sRegEncryptionPassword, null).ToString();
                debugMessage += "Encryption Password: " + sEncryptionPassword + "\n";
            }
            catch (System.Exception e)
            {
                MessageBox.Show("The SPAM Submission plug in has failed to load.\n" +
                    "Please contact support and tell them your reg keys need re-configuring\n",
                    "Error",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (bDebug)
            {
                MessageBox.Show(debugMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return true; 
        }

        private void blacklistSender(string sSenderEmailAddress)
        {
            //Now, let's add the sender to the autodelete rule
            olRuleList = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();

            foreach (Rule rule in olRuleList)
            {
                if (rule.Name.Equals(olRuleName))
                {
                    olRule = rule;
                    bSpamRuleExists = true;
                    break;
                }
            }

            if (!bSpamRuleExists)
            {
                //confirm that it doesn't exist and this isn't just a first run setting
                
                //if the rule still doesn't exist, we create it. 
                if (!bSpamRuleExists)
                {
                    olRule = olRuleList.Create(olRuleName, OlRuleType.olRuleReceive);
                    olRule.Conditions.SenderAddress.Address = new string[] { "1@2.3" };
                }
            }

            //then we check to see if the sender is in the bad list. If he's not,
            //we add him.
            bool bSenderInList = false;
            foreach (string s in olRule.Conditions.SenderAddress.Address)
            {
                if (s.Equals(sSenderEmailAddress))
                {
                    bSenderInList = true;
                    break;
                }
            }
            if (!bSenderInList)
            {
                //new
                string[] saBadAddresses = new string[olRule.Conditions.SenderAddress.Address.Length + 1];
                int i = 0;
                foreach (string s in olRule.Conditions.SenderAddress.Address)
                {
                    saBadAddresses[i] = s;
                    i++;
                }
                saBadAddresses[i] = sSenderEmailAddress;
                olRule.Conditions.SenderAddress.Address = saBadAddresses;
                olRule.Conditions.SenderAddress.Enabled = true;
            }
            olRule.Actions.DeletePermanently.Enabled = true;
            olRuleList.Save(true);
        }
    }
}
