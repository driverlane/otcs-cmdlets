using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using System.ComponentModel;

namespace cscmdlets
{
    [RunInstaller(true)]
    public class GetProcPSSnapIn01 : PSSnapIn
    {
        public override string Name
        {
            get
            {
                return "cscmdlets";
            }
        }

        public override string Vendor
        {
            get
            {
                return "Mark Farrall";
            }
        }

        public override string Description
        {
            get
            {
                return "Snap-in for a collection of cmdlets for managing OpenText Content Server.";
            }
        }
    }
}
