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
                Firewall.RemoveRemoteDesktopRule();

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

RemoteDesktopEnabler firewall [status|enable|on|disable|off]
status: Returns the current firewall status for private, domain and public networks. Enabled/Disabled
enable/on: Forces the remote desktop firewall rule to be enabled in Windows firewall. Use with care!
disable/off: Forces the remote desktop firewall rule to be disabled in Windows firewall. Use with care!

RemoteDesktopEnabler rdp [status|enable|on|disable|off]
status: Returns if remote desktop connections will be accepted by the computer
enable/on: Forces remote desktop connections to be accepted. Use with care!
disable/off: Forces remote desktop connections to be refused. Use with care!

Press ENTER to continue...");

            Console.ReadLine();
        }

        private void InterpretFirewallCommand(string command)
        {
            command = command.ToLower().Trim();

            if (command == "status")
            {
                ShowFirewallStatus();
                Console.ReadLine();
            }
            else if (command == "enable" || command == "on")
            {
                ForceFirewallRdpEnable();
            }
            else if (command == "disable" || command == "off")
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
                Console.ReadLine();
            }
            else if (command == "enable" || command == "on")
            {
                new Program().ForceRdpEnable();
            }
            else if (command == "disable" || command == "off")
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
                    Console.WriteLine();
                    Console.WriteLine();
                    Console.WriteLine("Press ENTER to disable firewall rule and close the program");
                    Console.ReadLine();
                    Firewall.RemoveRemoteDesktopRule();
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
            Firewall.RemoveRemoteDesktopRule();
            ShowFirewallStatus();
        }

        private void ForceFirewallRdpEnable()
        {
            Firewall.AddRemoteDesktopRule();
            ShowFirewallStatus();
        }

        private void ForceRdpEnable()
        {
            Rdp.SetRdpEnabled(true, true);
            ShowRdpStatus();
        }

        private void ForceRdpDisable()
        {
            Rdp.SetRdpEnabled(false, true);
            ShowRdpStatus();
        }

        private void TemporarelyEnableRemoteDesktop(int seconds)
        {
            Rdp.SetRdpEnabled(true);
            ShowRdpStatus();

            Console.WriteLine();

            Firewall.AddRemoteDesktopRule();
            ShowFirewallStatus();

            Console.WriteLine();

            System.Threading.Thread.Sleep(seconds * 1000);

            Rdp.SetRdpEnabled(false);
            ShowRdpStatus();

            Console.WriteLine();
        }

        private void ShowRdpStatus()
        {
            var status = Rdp.GetStatus();
            Console.WriteLine($"Remote Desktop connections {status.ToString()}");
        }

        private void ShowFirewallStatus()
        {
            var status = Firewall.Status(FirewallDomain.Private);
            Console.WriteLine($"Private network firewall is {status}");
            status = Firewall.Status(FirewallDomain.Domain);
            Console.WriteLine($"Domain network firewall is {status}");
            status = Firewall.Status(FirewallDomain.Public);
            Console.WriteLine($"Public network firewall is {status}");

            if (Firewall.RemoteDesktopRuleExists())
            {
                Console.WriteLine("Remote Desktop Firewall Rule is Activated");
            }
            else
            {
                Console.Write("Remote Desktop Firewall Rule is Deactivated");
            }
        }
    }
}

