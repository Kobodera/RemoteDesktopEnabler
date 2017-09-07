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
        readonly string FirewallRuleName = "RDPEnabler";
        readonly INetFwPolicy2 policyManager;

        public Firewall()
        {
            policyManager = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
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
            return policyManager.Rules.OfType<INetFwRule>().Where(x => x.Name == FirewallRuleName).Count() > 0;
        }

        private void AddCustomRemoteDesktopRule()
        {
            INetFwRule rule = (INetFwRule)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwRule"));
            rule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
            rule.Protocol = 6; //TCP
            rule.Enabled = true;
            rule.Name = FirewallRuleName;
            rule.InterfaceTypes = "All";
            rule.RemoteAddresses = "LocalSubnet";
            rule.LocalPorts = "3389";
            rule.RemotePorts = "3389";

            policyManager.Rules.Add(rule);
        }

        public static void AddRemoteDesktopRule()
        {
            new Firewall().AddCustomRemoteDesktopRule();
        }

        private void RemoveCustomRemoteDesktopRule()
        {
            if (RemoteDesktopFirewallRuleExists())
            {
                policyManager.Rules.Remove(FirewallRuleName);
            }
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
