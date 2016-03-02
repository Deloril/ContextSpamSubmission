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
        const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";
        //variable declarations XXX
        //the registry hive containing our address keys.
        string regPath = "HKEY_CURRENT_USER\\SOFTWARE\\InverseSoftware\\SpamSubmission\\";
        //the key containing the ticket 'voicemail' address. Emailing this address should result in
        //a ticket being created, with a reference to the SPAM sample.
        string regTicketAddress = "ticketEmail";
        //the key containing the address we submit the SPAM sample to.
        string regSubmitAddress = "spamEmail";
        //key that holds the zip password
        string regEncryptionPassword = "encryptionPassword";
        //string to store the registry key holding the debug value
        string regDebug = "debug";
        //key to hold the ticket voicemail address, once we get it.
        string emailTicketAddress = "";
        //key to hold the SPAM submission address, once we get it.
        string spamSubmitAddress = "";
        //string to store encryption password
        string encryptionPassword = "";
        //A string to hold the interesting items we want to report on in plaintext
        string metadata = "";
        //A string to hold the unique identifier of each message
        string uid = "";
        //boolean value (stored in reg) dictating wether or not we should show debugging messages
        bool debug = false;
        //an Outlook Rules Array to store all the current outlook rules.
        Rules olRuleList = null;
        //Single Rule instance
        Rule olRule = null;
        //A string for the rule name we will use. Registry?
        string olRuleName = "SPAMAutoDeleteList";

        
        
        public void submit()
        {
            Microsoft.Office.Interop.Outlook.Application outlookApp = Globals.ThisAddIn.Application;
            //submission zip password

            //Main logic, majority of program logic is below, in this method.
            Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
            {
                //item = variable storing what was right clicked on.
                object item = explorer.Selection[1];
                //if the item selected is a mail item, we know the user has done it right, let's proceed.
                if (item is MailItem)
                {
                    MailItem badMail = item as MailItem;
                    if (badMail != null)
                    {
                        //set the uid of this message
                        //after much deliberation, we'll initialize rand here,
                        //as it gives us the greatest chance of a non duplicated random number
                        string uid = Guid.NewGuid().ToString();
                        string strHeaders = "";
                        PropertyAccessor oPA = badMail.PropertyAccessor as PropertyAccessor;
                        const string PR_MAIL_HEADER_TAG = @"http://schemas.microsoft.com/mapi/proptag/0x007D001E";
                        try
                        {
                            strHeaders = (string)oPA.GetProperty(PR_MAIL_HEADER_TAG);
                        }
                        catch { }

                        //This will pull out the headers and such, and whack them into variables.
                        metadata = "To: " +  badMail.To + "\r\n";
                        metadata += "From: " +  badMail.SenderName + ": " + badMail.SenderEmailAddress + "\r\n";
                        metadata += "Subject: " + badMail.Subject + "\r\n";
                        metadata += "CC: " + badMail.CC + "\n\r";
                        metadata += "Companies Associated With Email: " + badMail.Companies + "\r\n";
                        metadata += "Email Creation Time: " + badMail.CreationTime + "\r\n";
                        metadata += "Delivery Report Requested: " +badMail.OriginatorDeliveryReportRequested + "\r\n";
                        metadata += "Received Time: " + badMail.ReceivedTime + "\r\n";
                        metadata += "Sent On: " + badMail.SentOn.ToString() + "\r\n";
                        metadata += "Size (kb): " + ((badMail.Size)/1024).ToString() + "\r\n";
                        metadata += "Headers: \r\n" + strHeaders + "\r\n";
                        metadata += "Plaintext Body: \r\n" + badMail.Body + "\r\n";

                        //This will create a mail item, and send it to the designated mailbox of a ticketing system.
                        Microsoft.Office.Interop.Outlook.MailItem ticketMail = (Microsoft.Office.Interop.Outlook.MailItem) outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        ticketMail.To = emailTicketAddress;
                        ticketMail.Subject = uid;
                        ticketMail.Body = metadata;
                        //ticketMail.Send();


                        //Save the badmail to disk, to then read back in in a compressed stream.
                        
                        //First, get temp path(checks the below in order):
                        //The path specified by the TMP environment variable.
                        //The path specified by the TEMP environment variable.
                        //The path specified by the USERPROFILE environment variable.
                        //The Windows directory.
                        string tempDir = Path.GetTempPath();
                        string badOnDisk = tempDir + uid + ".msg";
                        string badZipOnDisk = tempDir + uid + ".zip";

                        //Save it to disk
                        badMail.SaveAs(badOnDisk);

                        //Read in the email in the .zip format, with a password, write back to disk.
                        using (ZipFile zip = new ZipFile())
                        {
                            zip.Password = encryptionPassword;
                            //the "." specifies the directory structure inside the zip - . just means
                            //insert the attachment at the root, instead of nested in a replication of
                            //the systems %TMP% dir
                            zip.AddFile(badOnDisk, ".");
                            zip.Save(badZipOnDisk);
                        }

                        //Better get rid of the raw BadMail as soon as we're done with it

                        File.Delete(badOnDisk);

                        //This will create a mail item, and send it to a sample collection mailbox, with the badSample attached.
                        MailItem spamMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
                        spamMail.To = spamSubmitAddress;
                        spamMail.Subject = uid;
                        spamMail.Body = metadata;
                        spamMail.Attachments.Add(badZipOnDisk, OlAttachmentType.olByValue, 1, "SPAM Sample " + uid);
                        spamMail.Send();

                        //That's sent, let's delete the .zip on disk
                        File.Delete(badZipOnDisk);



                        //Now, let's add the sender to the autodelete rule
                        olRuleList = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();

                        bool ruleExists = false;
                        foreach (Rule rule in olRuleList)
                        {
                            if (rule.Name.Equals(olRuleName))
                            {
                                olRule = rule;
                                ruleExists = true;
                                break;
                            }
                        }

                        //if the rule doesn't exist, we create it. 
                        if (!ruleExists)
                        {
                            olRule = olRuleList.Create(olRuleName, OlRuleType.olRuleReceive);
                            olRule.Conditions.SenderAddress.Address = new string[] { "placeholder@ignoreme1337.ru" };
                        }
                        //then we check to see if the sender is in the bad list. If he's not,
                        //we add him.
                        bool inList = false;
                        foreach (string s in olRule.Conditions.SenderAddress.Address)
                        {
                            if (s.Equals(badMail.SenderEmailAddress)){
                                inList = true;
                                break;
                            }
                        }
                        if (!inList)
                        {
                            List<string> badAddressList = new List<string>();
                            foreach (string s in olRule.Conditions.SenderAddress.Address)
                            {
                                badAddressList.Add(s);

                            }
                            badAddressList.Add(badMail.SenderEmailAddress);
                            string[] badStrings = new string[badAddressList.Count];
                            int i = 0;
                            foreach (string s in badAddressList)
                            {
                                badStrings[i] = s;
                                i++;
                            }
                            foreach(string s in badStrings)
                            {
                                MessageBox.Show(s);
                            }
                            olRule.Conditions.SenderAddress.Address = badStrings;
                            olRule.Conditions.SenderAddress.Enabled = true;
                        }
                        olRule.Actions.DeletePermanently.Enabled = true;
                        olRuleList.Save(true);

                        //Finally, remove the dodgy email from outlook.
                        badMail.UnRead = false;
                        badMail.Save();
                        badMail.Delete();

                        //testing message, can likely remove this later. XXX
                        if (debug)
                        {
                            MessageBox.Show("You've submitted a SPAM sample.\r\n" +
                                metadata, "Thanks", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    //Not a mail item, need to decide how to handle this. Advise user they done goofed. XXX
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
            //debug string for testing
            string debugMessage= "";
            //let's grab our stuff from the registry
            try
            {
                debug = Registry.GetValue(regPath, regDebug, "false").ToString().ToLower().Equals("true");

                emailTicketAddress = Registry.GetValue(regPath, regTicketAddress, null).ToString();
                debugMessage += "Email Ticket Address: " + emailTicketAddress + "\n";

                spamSubmitAddress = Registry.GetValue(regPath, regSubmitAddress, null).ToString();
                debugMessage += "Spam Submit Address: " + spamSubmitAddress + "\n";

                encryptionPassword = Registry.GetValue(regPath, regEncryptionPassword, null).ToString();
                debugMessage += "Encryption Password: " + encryptionPassword + "\n";
            }
            catch (System.Exception e)
            {
                MessageBox.Show("The SPAM Submission plug in has failed to load.\n" +
                    "Please contact support and tell them your reg keys need re-configuring\n",
                    "Error",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            MessageBox.Show(debugMessage, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true; 
        }
    }
}
