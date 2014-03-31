using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace cscommandlets
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
        [Parameter(Mandatory = true, HelpMessage="e.g. http://server.domain/cws/")]
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
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory=false)]
        public Int64 TemplateID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            Node response = new Node();

            try
            {
                // create the connection object
                if (!Globals.Opened)
                {
                    response.Message = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
                    WriteObject(response);
                    return;
                }
                Connection connection = new Connection();

                // create the project workspace
                response = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Project);

                // if we've got a template ID then copy the config
                if (TemplateID > 0 && Convert.ToInt64(response) > 0)
                {
                    connection.UpdateProjectFromTemplate(Convert.ToInt64(response), TemplateID);
                }

                // write the output
                if (!String.IsNullOrEmpty(connection.ErrorMessage))
                {
                    response.Message = connection.ErrorMessage;
                    WriteObject(response);
                    return;
                }
                WriteObject(response);
            }
            catch (Exception e)
            {
                response.Message = e.Message;
                WriteObject(response);
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
        public Int64 ParentID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            Node response = new Node();

            try
            {

                // create the connection object
                if (!Globals.Opened)
                {
                    response.Message = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
                    WriteObject(response);
                    return;
                }
                Connection connection = new Connection();

                // create the folder
                response = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Folder);

                // write the output
                if (!String.IsNullOrEmpty(connection.ErrorMessage))
                {
                    response.Message = connection.ErrorMessage;
                    WriteObject(response);
                    return;
                }
                WriteObject(response);
            }
            catch (Exception e)
            {
                response.Message = e.Message;
                WriteObject(response);
            }

        }

    }

    [Cmdlet(VerbsCommon.Remove, "CSNode")]
    public class RemoveCSNode : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            String response = "";

            try
            {

                // create the connection object
                if (!Globals.Opened)
                {
                    response = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
                    WriteObject(response);
                    return;
                }
                Connection connection = new Connection();

                // create the folder
                response = connection.DeleteNode(NodeID);

                // write the output
                if (!String.IsNullOrEmpty(connection.ErrorMessage))
                {
                    response = connection.ErrorMessage;
                    WriteObject(response);
                    return;
                }
                WriteObject(response);
            }
            catch (Exception e)
            {
                response = e.Message;
                WriteObject(response);
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

    public class Node
    {
        public Int64 ID { get; set; }
        public String Message { get; set; }
    }
}
