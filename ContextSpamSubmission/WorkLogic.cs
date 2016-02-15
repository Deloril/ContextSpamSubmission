using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        String regPath = "HKEY_CURRENT_USER\\SOFTWARE\\ElementalSoftware\\";
        //the key containing the ticket 'voicemail' address. Emailing this address should result in
        //a ticket being created, with a reference to the SPAM sample.
        String regTicketAddress = "ticketsEmail";
        //the key containing the address we submit the SPAM sample to.
        String regSubmitAddress = "spamEmail";
        //key to hold the ticket voicemail address, once we get it.
        String emailTicketAddress = "";
        //key to hold the SPAM submission address, once we get it.
        String spamSubmitAddress = "";
        public void submit()
        {


            //read in addresses and persistent settings from registry. XXX
            //ticket voicemail address
            String temp = regPath + regTicketAddress;

            //safe place to send spam to

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
                        //testing message, can likely remove this later. XXX
                        MessageBox.Show("You've submitted something you think is SPAM!\r\n" +
                        badMail.Subject, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        //This will pull out the headers and such, and whack them into variables.


                        //This will create a mail item, and send it to the designated mailbox of a ticketing system.


                        //This will create a mail item, and send it to a sample collection mailbox, with the badSample attached.


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
            String debug = "";
            //let's grab our stuff from the registry
            try
            {
                emailTicketAddress = Registry.GetValue(regPath, regTicketAddress, null).ToString();
                debug += "Email Ticket Address: " + emailTicketAddress + "\n";

                spamSubmitAddress = Registry.GetValue(regPath, regSubmitAddress, null).ToString();
                debug += "Spam Submit Address: " + spamSubmitAddress + "\n";
            }
            catch (System.Exception e)
            {
                MessageBox.Show("The SPAM Submission plug in has failed to load.\n" +
                    "Please contact support and tell them your reg keys need re-configuring\n",
                    "Error",MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            MessageBox.Show(debug, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true; 
        }
    }
}
