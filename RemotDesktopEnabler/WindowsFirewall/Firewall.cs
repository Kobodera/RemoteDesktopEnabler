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
        readonly INetFwPolicy2 mgr = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

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
                fwCurrentProfileTypes = (NET_FW_PROFILE_TYPE2_)mgr.CurrentProfileTypes;
            }

            return (FirewallStatus)Convert.ToInt32(mgr.get_FirewallEnabled(fwCurrentProfileTypes));

        }
        internal static FirewallStatus Status(FirewallDomain? domain = null)
        {
            return new Firewall().FirewallStatus(domain);
        }

        private bool RemoteDesktopFirewallRuleExists()
        {
            return mgr.Rules.OfType<INetFwRule>().Where(x => x.LocalPorts == "3389").Count() > 0;
        }

        public static bool RemoteDesktopRuleExists()
        {
            return new Firewall().RemoteDesktopFirewallRuleExists();
        }

        private bool SetRemoteDesktopFirewallRule(bool enabled, bool forceChange)
        {
            bool allowChange = true;

            var rules = mgr.Rules.OfType<INetFwRule>().Where(x => x.LocalPorts == "3389");

            foreach (var rule in rules)
            {
                if (rule.Enabled == enabled)
                    allowChange = enabled || forceChange;
            }

            if (allowChange)
            {
                foreach (var rule in rules)
                {
                    rule.Enabled = enabled;
                }
            }

            return allowChange;
        }
        public static bool SetRemoteDesktopEnabled(bool enabled, bool forceChange = false)
        {
            return new Firewall().SetRemoteDesktopFirewallRule(enabled, forceChange);
        }
    }
}
