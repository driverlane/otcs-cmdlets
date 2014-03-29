using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace CGIPSTools
{

    [Cmdlet(VerbsCommon.Open, "CSConnection")]
    public class OpenCSConnectionCommand : Cmdlet
    {

        [Parameter(Mandatory = true)] 
        [ValidateNotNullOrEmpty]
        public String Username { get; set; }
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Password { get; set; }
        [Parameter(Mandatory = true, HelpMessage="e.g. http://mmglibrary.myintranet.local/les-services/")]
        [ValidateNotNullOrEmpty]
        public String ServicesDirectory { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Globals.Username = Username;
                Globals.Password = Password;
                Globals.ServicesDirectory = ServicesDirectory;

                Connection connection = new Connection();

                if (connection.ErrorMessage == null)
                {
                    WriteObject("Connection established");
                }
                else
                {
                    WriteObject(connection.ErrorMessage);
                }
            }
            catch (Exception e)
            {
                WriteObject(e.Message);
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSProjectWorkspace")]
    public class AddCSProjectWorkspaceCommand : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public Int32 ParentID { get; set; }
        [Parameter(Mandatory=false)]
        public Int32 TemplateID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            AGANode newNode = new AGANode();

            try
            {
                // create the connection object
                if (!Globals.Opened)
                {
                    newNode.Message = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
                    WriteObject(newNode);
                    return;
                }
                Connection connection = new Connection();

                // create the project workspace
                newNode = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Project);

                // if we've got a template ID then copy the config
                if (TemplateID > 0 && newNode.ID > 0)
                {
                    connection.UpdateProjectFromTemplate(newNode.ID, TemplateID);
                }

                // write the output
                if (newNode.ID == 0)
                {
                    newNode.Message = connection.ErrorMessage;
                    WriteObject(newNode);
                    return;
                }
                WriteObject(newNode);
            }
            catch (Exception e)
            {
                newNode.Message = e.Message;
                WriteObject(newNode);
                return;
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSFolder")]
    public class AddCSFolderCommand : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public Int32 ParentID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            AGANode newNode = new AGANode();

            try
            {

                // create the connection object
                if (!Globals.Opened)
                {
                    newNode.Message = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
                    WriteObject(newNode);
                    return;
                }
                Connection connection = new Connection();

                // create the folder
                newNode = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Folder);

                // write the output
                if (newNode.ID == 0)
                {
                    newNode.Message = connection.ErrorMessage;
                    WriteObject(newNode);
                    return;
                }
                WriteObject(newNode);
            }
            catch (Exception e)
            {
                newNode.Message = e.Message;
                WriteObject(newNode);
            }

        }

    }

    [Cmdlet(VerbsData.ConvertTo, "CGIEncryptedPassword")]
    public class ConvertToEncryptedPasswordCommand : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Password { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                WriteObject(Encryption.EncryptString(Password));
            }
            catch (Exception e)
            {
                WriteObject(e.Message);
            }
        }

    }

    public class AGANode
    {
        public Int32 ID { get; set; }
        public String NodeValue { get; set; }
        public String Message { get; set; }
    }
}
