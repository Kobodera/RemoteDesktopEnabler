using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using StringExtensions;

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

                Rdp.SetRdpEnabled(false1);

                Console.Write("Remote desktop connections disabled.");
            }
            return false;
        }

        private static void ShowUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("Enables remote desktop connections for a specified number of seconds.");
            Console.WriteLine("RemoteDesktopEnable <timeInSeconds>");
            Console.WriteLine("This example will enable remote connections for 60 seconds.");
            Console.WriteLine("RemoteDesktopEnabler 60");
            Console.WriteLine();
            Console.WriteLine("Force remote desktop connections to be enabled. Will not revert back to disabled.");
            Console.WriteLine("RemoteDesktopEnable enable");
            Console.WriteLine("Force remote desktop connections to be disabled. Will not revert back to enabled.");
            Console.WriteLine("RemoteDesktopEnable disable");
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
                if (args[0].ToLower() == "status")
                {
                    new Program().ShowRdpStatus();
                }

                if (args[0].ToLower() == "enable")
                {
                    new Program().ForceEnable();
                    return;
                }

                if (args[0].ToLower() == "disable")
                {
                    new Program().ForceDisable();
                    return;
                }

                if (!args[0].IsInt())
                {
                    ShowUsage();
                    return;
                }

                new Program().TemporarelyEnableRemoteDesktop(args[0].ToInt());

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.ReadLine();
            }
        }

        private void ForceEnable()
        {
            Console.WriteLine("Attempting to enable remote desktop connections by force.");
            Rdp.SetRdpEnabled(true, true);
            Console.WriteLine("Remote desktop connections enabled.");
        }

        private void ForceDisable()
        {
            Console.WriteLine("Attempting to disable remote desktop connections by force.");
            Rdp.SetRdpEnabled(false, true);
            Console.WriteLine("Remote desktop connections disabled.");
        }

        private void TemporarelyEnableRemoteDesktop(int seconds)
        {
            disableOnClose = true;

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
    }
}
