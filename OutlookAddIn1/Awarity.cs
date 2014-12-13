using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class Awarity
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += Application_ItemSend;
        }

        void Application_ItemSend(object Item, ref bool Cancel) {
            int score = 0;
            Cancel = false;
            if (Item is Microsoft.Office.Interop.Outlook.MailItem) {
                Microsoft.Office.Interop.Outlook.MailItem currentItem = Item as Microsoft.Office.Interop.Outlook.MailItem;
                // Check if recipient(s) is a saved contact{
                if (!currentItem.Recipients.ResolveAll()) {
                    score++;
                }

                // Check if recipient is a blacklisted domain (e.g. google safeapi)

                // Check if there is an attachment
                if (currentItem.Attachments.Count >= 0) {
                    score++;
                }
                
                /* ASK THE USER */
                if (score>=1){
                    string message = "Is there really no confidential stuff in the mail? Do you know the recipient(s)?";
                    string cap = "Alert!";

                    var res = MessageBox.Show(message, cap, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (res == DialogResult.No) {
                        Cancel = true;
                    }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
