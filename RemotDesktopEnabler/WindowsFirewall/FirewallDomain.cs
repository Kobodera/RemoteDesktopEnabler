using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RemotDesktopEnabler
{
    internal enum FirewallDomain
    {
         Domain = 0x0001,
         Private = 0x0002,
         Public = 0x0004,
         All = 0x7FFFFFFF
    }
}
