using System;

namespace cscmdlets
{

    internal class Globals
    {
        internal static String Username;
        internal static String Password;
        internal static String ServicesDirectory;
        internal static String RestUrl;
        internal static Boolean SoapConnectionOpened;
        internal static Boolean RestConnectionOpened;

        internal enum PhysicalItemTypes
        {
            PhysicalItem,
            PhysicalItemContainer,
            PhysicalItemBox
        }

        internal enum ObjectType
        {
            Folder,
            Project
        };

    }

}
