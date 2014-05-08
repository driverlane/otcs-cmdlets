using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace cscommandlets
{

    #region encryption

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
                ErrorRecord err = new ErrorRecord(e, "ConvertToEncryptedPasswordCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }
    }

    #endregion

    #region Connection

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

        Connection connection;

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Globals.Username = Username;
                Globals.Password = Password;
                Globals.ServicesDirectory = ServicesDirectory;

                connection = new Connection();

                WriteObject("Connection established");
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "OpenCSConnectionCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
                return;
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "OpenCSConnectionCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region document management webservice

    [Cmdlet(VerbsCommon.Add, "CSProjectWorkspace")]
    public class AddCSProjectWorkspaceCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory=false)]
        public Int64 TemplateID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSProjectWorkspaceCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                Boolean addTemplate = true;

                // create the project workspace
                Int64 response = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Project);

                // throw any nonterminating errors
                foreach (Exception ex in connection.NonTerminatingExceptions)
                {
                    if (ex.Message.EndsWith("already exists.")) addTemplate = false;
                    ErrorRecord err = new ErrorRecord(ex, "AddCSProjectWorkspaceCommand", ErrorCategory.NotSpecified, this);
                    WriteError(err);
                }


                // if we've got a template ID then copy the config
                if (TemplateID > 0 && Convert.ToInt64(response) > 0 && addTemplate)
                {
                    connection.UpdateProjectFromTemplate(Convert.ToInt64(response), TemplateID);
                }

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSProjectWorkspaceCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
                return;
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "OpenCSConnectionCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSFolder")]
    public class AddCSFolderCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSFolderCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the folder
                Int64 response = connection.CreateContainer(Name, ParentID, Connection.ObjectType.Folder);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSFolderCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSFolderCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Remove, "CSNode")]
    public class RemoveCSNodeCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSNodeCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the folder
                String response = connection.DeleteNode(NodeID);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSNodeCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSNodeCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region Member webservices

    [Cmdlet(VerbsCommon.Add, "CSUser")]
    public class AddCSUserCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Login { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 DepartmentGroupID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Password { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String FirstName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String MiddleName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LastName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Email { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Fax { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String OfficeLocation { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Phone { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Title { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? LoginEnabled { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? PublicAccessEnabled { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? CreateUpdateUsers { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? CreateUpdateGroups { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? CanAdministerUsers { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? CanAdministerSystem { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                // set the privileges to default if null
                Boolean blnLoginEnabled = LoginEnabled ?? true;
                Boolean blnPublicAccessEnabled = PublicAccessEnabled ?? true;
                Boolean blnCreateUpdateUsers = CreateUpdateUsers ?? false;
                Boolean blnCreateUpdateGroups = CreateUpdateGroups ?? false;
                Boolean blnCanAdministerUsers = CanAdministerUsers ?? false;
                Boolean blnCanAdministerSystem = CanAdministerSystem ?? false;

                // create the user
                Int64 response = connection.CreateUser(Login, DepartmentGroupID, Password, FirstName, MiddleName, LastName, Email, Fax, OfficeLocation, 
                    Phone, Title, blnLoginEnabled, blnPublicAccessEnabled, blnCreateUpdateUsers, blnCreateUpdateGroups, blnCanAdministerUsers, blnCanAdministerSystem);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Remove, "CSUser")]
    public class RemoveCSUserCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 UserID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

                // create the user
                String response = connection.DeleteUser(UserID);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSUserCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Get, "CSUserIDByLogin")]
    public class GetCSUserIDByLoginCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Login { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSUserIDByLoginCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // get the username
                Int64 UserID = connection.GetUserIDByLoginName(Login);

                // write the output
                WriteObject(UserID);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSUserIDByLoginCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSUserIDByLoginCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region Classifications webservice

    [Cmdlet(VerbsCommon.Add, "CSClassifications")]
    public class AddCSClassificationsCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64[] ClassificationIDs { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSClassificationsCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                String response = "";

                // add the classifications
                Boolean success = connection.AddClassifications(NodeID, ClassificationIDs);
                if (success)
                {
                    response = "Classifications applied";
                }
                else
                {
                    // i'm assuming we never get to this because an exception is thrown, but would need more testing
                    response = "Classifications not applied";
                }

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSClassificationsCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSClassificationsCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region Records management webservice

    [Cmdlet(VerbsCommon.Add, "CSRMClassification")]
    public class AddCSRMClassificationCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 RMClassificationID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSRMClassificationCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                String response = "";

                // add the RM classification
                Boolean success = connection.AddRMClassification(NodeID, RMClassificationID);
                if (success)
                {
                    response = "RM classification applied";
                }
                else
                {
                    // i'm assuming we never get to this because an exception is thrown, but would need more testing
                    response = "RM classification not applied";
                }

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSRMClassificationCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSRMClassificationCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Set, "CSFinaliseRecord")]
    public class SetCSFinaliseRecordCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSFinaliseRecordCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                String response = "";

                // finalise the item
                Boolean success = connection.FinaliseRecord(NodeID);
                if (success)
                {
                    response = "Item finalised";
                }
                else
                {
                    // i'm assuming we never get to this because an exception is thrown, but would need more testing
                    response = "Item not finalised";
                }

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSFinaliseRecordCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSFinaliseRecordCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region Physical objects webservice

    [Cmdlet(VerbsCommon.Set, "CSPhysObjToBox")]
    public class SetCSPhysObjToBox : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ItemID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 BoxID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? UpdateLocation { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? UpdateRSI { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean? UpdateStatus { get; set; }

        Connection connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.ConnectionOpened)
                {
                    ThrowTerminatingError(Errors.ConnectionMissing(this));
                    return;
                }
                connection = new Connection();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetPhysObjToBox", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                String response = "";

                // convert any null booleans - defaults to false
                Boolean blnUpdateLocation = UpdateLocation ?? false;
                Boolean blnUpdateRSI = UpdateRSI ?? false;
                Boolean blnUpdateStatus = UpdateStatus ?? false;

                // assign to the box
                Boolean success = connection.AssignToBox(ItemID, BoxID, blnUpdateLocation, blnUpdateRSI, blnUpdateStatus);
                if (success)
                {
                    response = "Assigned to box";
                }
                else
                {
                    // i'm assuming we never get to this because an exception is thrown, but would need more testing
                    response = "Not assigned to box";
                }

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetPhysObjToBox", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();

            try
            {
                connection.CloseClients();
            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetPhysObjToBox", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    internal class Errors
    {

        internal static ErrorRecord ConnectionMissing(Object Object)
        {
            String msg = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
            Exception exception = new Exception(msg);
            ErrorRecord err = new ErrorRecord(exception, "ConnectionMissing", ErrorCategory.ResourceUnavailable, Object);
            return err;
        }

    }
}
