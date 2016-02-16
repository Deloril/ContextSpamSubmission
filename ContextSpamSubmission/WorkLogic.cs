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
        //variable declarations XXX
        //the registry hive containing our address keys.
        string regPath = "HKEY_CURRENT_USER\\SOFTWARE\\ElementalSoftware\\SpamSubmission\\";
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
        bool debug;
        
        public void submit()
        {


            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
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
                        Random rand = new Random();
                        uid = rand.Next(100000000, 999999999).ToString();
                        uid += rand.Next(100000000, 999999999).ToString();
                        uid += rand.Next(100000000, 999999999).ToString();

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
                        metadata += "Plaintext Body: \r\n" + badMail.Body + "\r\n";

                        //This will create a mail item, and send it to the designated mailbox of a ticketing system.
                        Microsoft.Office.Interop.Outlook.MailItem ticketMail = (Microsoft.Office.Interop.Outlook.MailItem) outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        ticketMail.To = emailTicketAddress;
                        ticketMail.Subject = uid;
                        ticketMail.Body = metadata;
                        ticketMail.Send();

                        MemoryStream ms = new MemoryStream();
                        using (ZipArchive zipper = new ZipArchive(ms))
                        {

                        }

                        
                        //Attachment badAttach = new System.Net.Mail.Attachment(badMail, System.Net.Mime.MediaTypeNames.Application.Octet);


                        //This will create a mail item, and send it to a sample collection mailbox, with the badSample attached.
                        Microsoft.Office.Interop.Outlook.MailItem spamMail = (Microsoft.Office.Interop.Outlook.MailItem)outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                        spamMail.To = spamSubmitAddress;
                        spamMail.Subject = uid;
                        spamMail.Body = metadata;
                        //spamMail.a
                        spamMail.Send();

                        //testing message, can likely remove this later. XXX
                        if (debug)
                        {
                            MessageBox.Show("You've submitted something you think is SPAM!\r\n" +
                        metadata, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    //Not a mail item, need to decide how to handle this. Advise user they done goofed. XXX
                    else
                    {
                        MessageBox.Show("Not an email!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
