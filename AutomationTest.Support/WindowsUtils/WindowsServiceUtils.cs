using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;

namespace AutomationUtils.WindowsUtls
{
    public class WindowsServiceUtils
    {
        public static string QueryWindowsServiceState(string machineName, string serviceName, string UserName, string Password)
        {
            string StateString = "";
           
            ConnectionOptions myConnection = new ConnectionOptions();

            myConnection.Username = UserName;
            myConnection.Password = Password;
            string providerPath = @"\\" + machineName + @"\root\cimv2";

            ManagementScope myScope = new ManagementScope(providerPath, myConnection);
            myScope.Connect();

            SelectQuery query = new SelectQuery("select * from Win32_Service where name = '" + serviceName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(myScope, query);
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject service in collection)
            {
                StateString = service.GetPropertyValue("State").ToString();
            }

            return StateString;
        }

        public static int StartWindowsService(string machineName, string serviceName, string UserName, string Password)
        {            
            ConnectionOptions myConnection = new ConnectionOptions();

            myConnection.Username = UserName;
            myConnection.Password = Password;
            string providerPath = @"\\" + machineName + @"\root\cimv2";

            ManagementScope myScope = new ManagementScope(providerPath, myConnection);
            myScope.Connect();

            SelectQuery query = new SelectQuery("select * from Win32_Service where name = '" + serviceName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(myScope, query);
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject service in collection)
            {
                service.InvokeMethod("StartService", null);
            }
            return 0;
        }

        public static int StopWindowsService(string machineName, string serviceName, string UserName, string Password)
        { 
            ConnectionOptions myConnection = new ConnectionOptions();

            myConnection.Username = UserName;
            myConnection.Password = Password;
            string providerPath = @"\\" + machineName + @"\root\cimv2";

            ManagementScope myScope = new ManagementScope(providerPath, myConnection);
            myScope.Connect();

            SelectQuery query = new SelectQuery("select * from Win32_Service where name = '" + serviceName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(myScope, query);
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject service in collection)
            {
                service.InvokeMethod("StopService", null);
            }
            return 0;
        }

        public static int RestartWindowsService(string machineName, string serviceName, string UserName, string Password)
        {            
            ConnectionOptions myConnection = new ConnectionOptions();

            myConnection.Username = UserName;
            myConnection.Password = Password;
            string providerPath = @"\\" + machineName + @"\root\cimv2";

            ManagementScope myScope = new ManagementScope(providerPath, myConnection);
            myScope.Connect();

            SelectQuery query = new SelectQuery("select * from Win32_Service where name = '" + serviceName + "'");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(myScope, query);
            ManagementObjectCollection collection = searcher.Get();

            foreach (ManagementObject service in collection)
            {
                if (service.GetPropertyValue("State").ToString().ToLower().Equals("running"))
                {
                    service.InvokeMethod("StopService", null);
                }

                service.InvokeMethod("StartService", null);
            }
            return 0;
        }
    }
}
