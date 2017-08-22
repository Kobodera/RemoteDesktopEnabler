using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StringExtensions;

namespace RemotDesktopEnabler
{
    internal class Rdp
    {
        private bool SetRdpRegistryValue(int value, bool forceChange)
        {
            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
            var rdpKey = key.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server", true);
            int currentValue = rdpKey.GetValue("fDenyTSConnections").ToString().ToInt(-1);

            //Value was not found do not proceed with change.
            if (currentValue == -1)
            {
                return false;
            }
            else if (value == 1 && currentValue == 1 && !forceChange)
            {
                Console.WriteLine("RDP is already disabled. No changes will be made.");
                return false;
            }
            else if (value == 0 && currentValue == 0 && !forceChange)
            {
                Console.WriteLine("RDP is already enabled. No changes will be made.");
            }
            else
            {
                rdpKey.SetValue("fDenyTSConnections", value);
            }

            return true;
        }

        internal static bool SetRdpEnabled(bool enabled, bool forceChange = false)
        {
            if (enabled)
            {
                return new Rdp().SetRdpRegistryValue(0, forceChange);
            }
            else
            {
                return new Rdp().SetRdpRegistryValue(1, forceChange);
            }
        }

        internal static RdpStatus GetStatus()
        {
            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
            var rdpKey = key.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Terminal Server", true);
            int currentValue = rdpKey.GetValue("fDenyTSConnections").ToString().ToInt(-1);

            return (RdpStatus)currentValue;
        }
    }
}
