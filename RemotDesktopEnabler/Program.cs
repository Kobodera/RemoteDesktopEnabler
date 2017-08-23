using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using StringExtensions;
using RemotDesktopEnabler.WindowsFirewall;
using RemotDesktopEnabler.RemoteDesktop;

namespace RemotDesktopEnabler
{
    /// <summary>
    /// This program enables remote desktop connections to a computer for a limited time. After the time runs out
    /// remote desktop connections are yet again disabled. If remote desktop connections are enabled before the program runs
    /// no changes will be made
    /// </summary>
    class Program
    {
        private static bool disableOnClose = false;

        static ConsoleEventDelegate handler;

        private delegate bool ConsoleEventDelegate(int eventType);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);

        static bool ConsoleEventCallback(int eventType)
        {
            //Console is being forcefully closed. Make sure that remote desktop is disabled
            if (eventType == 2 && disableOnClose)
            {
                Console.WriteLine("Attempting to disable remote desktop connections.");

                Rdp.SetRdpEnabled(false);

                Console.Write("Remote desktop connections disabled.");
            }
            return false;
        }

        private static void ShowUsage()
        {
            Console.WriteLine(@"
This program temporary allows Remote Desktop connections to be established.

It can force firewall and remote desktop rules to be either enabled or disabled.

Unless forcing a rule change systems with enabled rules will not be modified by the program.

Usage:

RemoteDesktopEnabler [Time in seconds]
Allows remote desktop connection for [Time in seconds] seconds.
Ex. RemoteDesktopEnabler 10
Allows Remote Desktop connections for 10 seconds.

RemoteDesktopEnabler firewall [exists|status|enable|disable]
exists: Returns if the Remote Desktop rule exists in Windows firewall. 
        Ex. Windows 10 Home edition does not allow remote connections
status: Returns the current firewall status for private, domain and public networks. Enabled/Disabled
enable: Forces the remote desktop firewall rule to be enabled in Windows firewall. Use with care!
disable: Forces the remote desktop firewall rule to be disabled in Windows firewall. Use with care!

RemoteDesktopEnabler rdp [status|enable|disable]
status: Returns if remote desktop connections will be accepted by the computer
enable: Forces remote desktop connections to be accepted. Use with care!
disable: Forces remote desktop connections to be refused. Use with care!

Press ENTER to continue...");

            Console.ReadLine();
        }

        private void InterpretFirewallCommand(string command)
        {
            command = command.ToLower().Trim();

            if (command == "exists")
            {
                ShowFirewallRdpRuleExists();
            }
            else if (command == "status")
            {
                ShowFirewallStatus();
            }
            else if (command == "enable")
            {
                ForceFirewallRdpEnable();
            }
            else if (command == "disable")
            {
                ForceFirewallRdpDisable();
            }
            else
            {
                ShowUsage();
            }
        }

        private void InterpretRdpCommand(string command)
        {
            command = command.ToLower().Trim();

            if (command == "status")
            {
                new Program().ShowRdpStatus();
            }
            else if (command == "enable")
            {
                new Program().ForceRdpEnable();
            }
            else if (command == "disable")
            {
                new Program().ForceRdpDisable();
            }
            else
            {
                ShowUsage();
            }
        }

        static void Main(string[] args)
        {
            if (args.Count() == 0)
            {
                ShowUsage();
                return;
            }

            handler = new ConsoleEventDelegate(ConsoleEventCallback);
            SetConsoleCtrlHandler(handler, true);

            try
            {
                if (args[0].ToLower() == "firewall")
                {
                    if (args.Length == 2)
                    {
                        new Program().InterpretFirewallCommand(args[1]);
                    }
                    else
                    {
                        ShowUsage();
                    }
                }
                else if (args[0].ToLower() == "rdp")
                {
                    if (args.Length == 2)
                    {
                        new Program().InterpretRdpCommand(args[1]);
                    }
                    else
                    {
                        ShowUsage();
                    }
                }
                else if (args[0].IsInt())
                {
                    new Program().TemporarelyEnableRemoteDesktop(args[0].ToInt());
                }
                else
                {
                    ShowUsage();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        private void ForceFirewallRdpDisable()
        {
            if (Firewall.RemoteDesktopRuleExists())
            {
                Console.WriteLine("Forcing Remote Desktop Firewall rule disabled.");
                Firewall.SetRemoteDesktopEnabled(false, true);
            }
        }

        private void ForceFirewallRdpEnable()
        {
            if (Firewall.RemoteDesktopRuleExists())
            {
                Console.WriteLine("Forcing Remote Desktop Firewall rule enabled.");
                Firewall.SetRemoteDesktopEnabled(true, true);
            }
        }

        private void ForceRdpEnable()
        {
            Console.WriteLine("Attempting to enable remote desktop connections by force.");
            Rdp.SetRdpEnabled(true, true);
            Console.WriteLine("Remote desktop connections enabled.");
        }

        private void ForceRdpDisable()
        {
            Console.WriteLine("Attempting to disable remote desktop connections by force.");
            Rdp.SetRdpEnabled(false, true);
            Console.WriteLine("Remote desktop connections disabled.");
        }

        private bool EnableRdpInFirewall()
        {
            if (!Firewall.RemoteDesktopRuleExists())
            {
                Console.WriteLine("No remote desktop rule found. No modification to firewall will be done.");
                return false;
            }

            Console.WriteLine("Attempting to enable Remote Desktop in Wondows firewall.");
            bool fwRulesChanged = Firewall.SetRemoteDesktopEnabled(true);
            if (fwRulesChanged)
            {
                Console.WriteLine("Firewall rules enabled");
            }
            else
            {
                Console.WriteLine("Remote Desktop firewall rules already enabled.");
            }

            return fwRulesChanged;
        }

        private void DisableRdpInFirewall(bool fwRulesChanged)
        {
            if (fwRulesChanged)
            {
                Console.WriteLine("Attempting to disable remote desktop firewall rules.");
                Firewall.SetRemoteDesktopEnabled(false);
                Console.WriteLine("Remote desktop firewall rules disabled.");
            }
        }

        private void TemporarelyEnableRemoteDesktop(int seconds)
        {

            disableOnClose = true;

            bool fwRulesChanged = EnableRdpInFirewall();

            Console.WriteLine("Attempting to enable remote desktop connections.");
            
            if (Rdp.SetRdpEnabled(true))
            {

                Console.WriteLine($"Remote desktop connections enabled for {seconds} seconds.");
            }
            else
            {
                Console.ReadLine();
                return;
            }

            System.Threading.Thread.Sleep(seconds * 1000);

            DisableRdpInFirewall(fwRulesChanged);

            Console.WriteLine("Attempting to disable remote desktop connections.");

            if (Rdp.SetRdpEnabled(false))
            {
                Console.Write("Remote desktop connections disabled.");
            }
            else
            {
                Console.ReadLine();
                return;
            }
        }

        private void ShowRdpStatus()
        {
            var status = Rdp.GetStatus();
            Console.WriteLine($"Remote Desktop connections {status.ToString()}");
            Console.ReadLine();
        }

        private void ShowFirewallStatus()
        {
            var status = Firewall.Status(FirewallDomain.Private);
            Console.WriteLine($"Private network firewall is {status}");
            status = Firewall.Status(FirewallDomain.Domain);
            Console.WriteLine($"Domain network firewall is {status}");
            status = Firewall.Status(FirewallDomain.Public);
            Console.WriteLine($"Public network firewall is {status}");
            Console.ReadLine();
        }

        private void ShowFirewallRdpRuleExists()
        {
            Console.WriteLine($"Remote desktop rules exists: {Firewall.RemoteDesktopRuleExists()}");
            Console.ReadLine();
        }

    }
}
