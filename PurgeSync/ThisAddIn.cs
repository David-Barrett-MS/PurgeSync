/*
  DISCLAIMER:
THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;

namespace PurgeSync
{
    public partial class ThisAddIn
    {
        private TimeSpan _purgeTimeSpan=new TimeSpan(2,0,0,0); // Default to 2 days purge period

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ReadPurgeTimeSpan();
            PurgeSyncFolders();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void PurgeSyncFolders()
        {
            // Purge the Sync Issues folder/sub-folders

            try
            {
                // Obtain reference to Sync Issues folder
                Outlook.Folder oSyncFolder = (Outlook.Folder)this.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSyncIssues);

                // Now process each subfolder
                foreach (Outlook.Folder oFolder in oSyncFolder.Folders)
                {
                    PurgeFolder(oFolder);
                }
                // Finally process the Sync Issues folder itself
                PurgeFolder(oSyncFolder);
            }
            catch { }
        }

        private void PurgeFolder(Outlook.Folder folder)
        {
            // Purge the specified folder of any items older than the given time span

            // Obtain a reference to the items in the folder
            Outlook.Items oItems = folder.Items;

            // Create a restriction so that we are only looking at items to be purged
            string sRestrict = "";
            try
            {
                DateTime oDeleteBefore = DateTime.Now.Subtract(_purgeTimeSpan);
                sRestrict = "[CreationTime] < '" + oDeleteBefore.ToShortDateString() + oDeleteBefore.ToString(" h:mm tt") + "'";
            }
            catch
            {
                // Failed to build restriction, do not continue
                return;
            }
            try
            {
                oItems = oItems.Restrict(sRestrict);
            }
            catch
            {
                // Failed on restriction, so do not delete anything
                return;
            }

            // Now delete the items
            for (int i = oItems.Count; i > 0; i--)
            {
                try
                {
                    oItems[i].Delete();
                }
                catch { }
            }
        }

        private void ReadPurgeTimeSpan()
        {
            // Check for registry setting to override default timespan
            RegistryKey oRegKey=null;
            try
            {
                oRegKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\PurgeSync");
                if (oRegKey == null)
                    oRegKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\PurgeSync");

                if (oRegKey == null)
                    return; // We haven't got a registry setting, so quit
            }
            catch
            {
                // Unexpected error, so quit
                return;
            }


            // Read the registry setting
            try
            {
                int iTimeSpan = (int)oRegKey.GetValue("PurgePeriod");
                _purgeTimeSpan = new TimeSpan(iTimeSpan, 0, 0, 0);
                oRegKey.Close();
            }
            catch { }
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
