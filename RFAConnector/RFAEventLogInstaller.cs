using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFAConnector
{
    [RunInstaller(true)]
    public class RFAEventLogInstaller : Installer
    {
        private EventLogInstaller myEventLogInstaller;

        public RFAEventLogInstaller()
        {
            // Create an instance of an EventLogInstaller.
            myEventLogInstaller = new EventLogInstaller();

            // Set the source name of the event log.
            myEventLogInstaller.Source = "RFAConnectorSource";

            // Set the event log that the source writes entries to.
            myEventLogInstaller.Log = "RFAConnectorLog";

            // Add myEventLogInstaller to the Installer collection.
            Installers.Add(myEventLogInstaller);
        }

        public static void Main()
        {
            RFAEventLogInstaller myInstaller = new RFAEventLogInstaller();
        }
    }
}
