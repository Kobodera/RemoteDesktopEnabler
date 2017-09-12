using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetFwTypeLib;

namespace RemotDesktopEnabler.WindowsFirewall
{
    internal class Firewall
    {
        readonly string FirewallRuleName = Properties.Settings.Default.FirewallRuleName;
        readonly INetFwPolicy2 policyManager;
        INetFwMgr manager;
        INetFwProfile profile;
        INetFwOpenPorts openPorts;

        public Firewall()
        {
            policyManager = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            manager = (INetFwMgr)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwMgr"));
            profile = manager.LocalPolicy.CurrentProfile;
            openPorts = profile.GloballyOpenPorts;
        }

        private FirewallStatus FirewallStatus(FirewallDomain? domain)
        {
            // Gets the current firewall profile (domain, public, private, etc.)
            NET_FW_PROFILE_TYPE2_ fwCurrentProfileTypes;

            if (domain.HasValue)
            {
                fwCurrentProfileTypes = (NET_FW_PROFILE_TYPE2_)domain;
            }
            else
            {
                fwCurrentProfileTypes = (NET_FW_PROFILE_TYPE2_)policyManager.CurrentProfileTypes;
            }

            return (FirewallStatus)Convert.ToInt32(policyManager.get_FirewallEnabled(fwCurrentProfileTypes));

        }
        internal static FirewallStatus Status(FirewallDomain? domain = null)
        {
            return new Firewall().FirewallStatus(domain);
        }

        private bool RemoteDesktopFirewallRuleExists()
        {
            return policyManager.Rules.OfType<INetFwRule>().Where(x => x.Name.StartsWith(FirewallRuleName)).Count() > 0;
        }

        private void OpenPort(string name, int port, NET_FW_IP_PROTOCOL_ protocol, NET_FW_SCOPE_ scope)
        {
            if (openPorts.OfType<INetFwOpenPort>().Where(x => x.Name == name).Count() == 0)
            {
                INetFwOpenPort openPort = (INetFwOpenPort)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwOpenPort"));
                openPort.Port = port;
                openPort.Protocol = protocol;
                openPort.Scope = scope;
                openPort.Name = name;

                openPorts.Add(openPort);
            }
        }

        private void AddCustomRemoteDesktopRule()
        {
            OpenPort($"{FirewallRuleName}_TCP", Properties.Settings.Default.RdpTcpPort, NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP, (NET_FW_SCOPE_)Properties.Settings.Default.RdpScope);
            OpenPort($"{FirewallRuleName}_UDP", Properties.Settings.Default.RdpUdpPort, NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_UDP, (NET_FW_SCOPE_)Properties.Settings.Default.RdpScope);
        }

        public static void AddRemoteDesktopRule()
        {
            new Firewall().AddCustomRemoteDesktopRule();
        }

        private void RemoveCustomRemoteDesktopRule()
        {
            openPorts.Remove(Properties.Settings.Default.RdpTcpPort, NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_TCP);
            openPorts.Remove(Properties.Settings.Default.RdpUdpPort, NET_FW_IP_PROTOCOL_.NET_FW_IP_PROTOCOL_UDP);

            policyManager.Rules.Remove(FirewallRuleName);
        }

        public static void RemoveRemoteDesktopRule()
        {
            new Firewall().RemoveCustomRemoteDesktopRule();
        }

        public static bool RemoteDesktopRuleExists()
        {
            return new Firewall().RemoteDesktopFirewallRuleExists();
        }
    }
}
