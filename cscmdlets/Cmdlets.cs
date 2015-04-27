using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;

namespace cscmdlets
{

    #region Encryption

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

    [Cmdlet(VerbsCommon.Open, "CSConnections")]
    public class OpenCSConnectionsCommand : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Username { get; set; }
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Password { get; set; }
        [Parameter(Mandatory = true, HelpMessage = "e.g. http://server.domain/cws/")]
        [ValidateNotNullOrEmpty]
        public String ServicesDirectory { get; set; }
        [Parameter(Mandatory = true, HelpMessage = "e.g. http://server.domain/otcs/cs.exe/api/v1/")]
        [ValidateNotNullOrEmpty]
        public String Url { get; set; }

        SoapApi connection;

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Globals.Username = Username;
                Globals.Password = Password;
                Globals.ServicesDirectory = ServicesDirectory;
                Globals.RestUrl = Url;

                // try and open a SOAP connection
                connection = new SoapApi();
                Globals.SoapConnectionOpened = true;

                // try and open a REST connection
                RestApi.CheckConnection();
                if (Globals.SoapConnectionOpened && Globals.RestConnectionOpened)
                    WriteObject("Connections established");
                else
                    WriteObject("Connections NOT established. ERROR: not specified.");

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "OpenCSConnectionsCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
                return;
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

    }

    [Cmdlet(VerbsCommon.Open, "CSConnectionSOAP")]
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

        SoapApi connection;

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Globals.Username = Username;
                Globals.Password = Password;
                Globals.ServicesDirectory = ServicesDirectory;

                connection = new SoapApi();
                Globals.SoapConnectionOpened = true;

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

    [Cmdlet(VerbsCommon.Open, "CSConnectionREST")]
    public class OpenCSConnectionRestCommand : Cmdlet
    {

        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Username { get; set; }
        [Parameter(Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public String Password { get; set; }
        [Parameter(Mandatory = true, HelpMessage = "e.g. http://server.domain/otcs/cs.exe/api/v1/")]
        [ValidateNotNullOrEmpty]
        public String Url { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Globals.Username = Username;
                Globals.Password = Password;
                Globals.RestUrl = Url;

                if (RestApi.CheckConnection())
                    WriteObject("Connection established");
                else
                    WriteObject("Connection NOT established. ERROR: not specified.");

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "OpenCSConnectionRestCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
                return;
            }
        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

    }

    #endregion

    #region Core commands

    [Cmdlet(VerbsCommon.Add, "CSFolder")]
    public class AddCSFolderCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                Int64 response = connection.CreateContainer(Name, ParentID, Globals.ObjectType.Folder);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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

    [Cmdlet(VerbsCommon.Add, "CSDocument")]
    public class AddCSDocumentCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Document { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSDocumentCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the folder
                Int64 response = connection.CreateSimpleDoc(Name, ParentID, Document);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSDocumentCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Remove, "CSNode")]
    public class RemoveCSNodeCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                RemoveNode(NodeID);
            }
            catch (Exception e)
            {
                String response = String.Format("{0} - NOT deleted. Error - {1}", NodeID, e.Message);
                WriteObject(response);
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

        internal void RemoveNode(Int64 thisNode)
        {

            try
            {

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            RemoveNode(child);
                        }
                    }
                }

                // remove the parent object
                String response;
                try
                {
                    connection.DeleteNode(thisNode);
                    response = String.Format("{0} - deleted", thisNode);
                }
                catch (Exception e)
                {
                    response = String.Format("{0} - NOT deleted", thisNode, e.Message);
                }
                WriteObject(response);

            }
            catch (Exception e)
            {
                String message = String.Format("{0} - NOT deleted. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }
    }

    #endregion

    #region Projects

    [Cmdlet(VerbsCommon.Add, "CSProjectWorkspace")]
    public class AddCSProjectWorkspaceCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = false)]
        public Int64 TemplateID { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                Int64 response = connection.CreateContainer(Name, ParentID, Globals.ObjectType.Project);

                // ignore any already exists errors
                foreach (Exception ex in connection.NonTerminatingExceptions)
                {
                    if (ex.Message.EndsWith("already exists."))
                    {
                        addTemplate = false;
                    }
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
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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

    #region Cats and atts

    [Cmdlet(VerbsCommon.Get, "CSCategories")]
    public class GetCSCategoriesCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public SwitchParameter ShowKey
        {
            get { return showKey; }
            set { showKey = value; }
        }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        SoapApi connection;
        Boolean showKey;
        Boolean recurse;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSCategoriesCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                ListCategories(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Categories NOT returned. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "GetCSCategoriesCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void ListCategories(Int64 thisNode)
        {
            try
            {

                // copy the category and add the log entry
                List<String> response = connection.ListNodeCategories(thisNode, showKey);
                List<String> cats = new List<string>();
                foreach (String cat in response)
                {
                    cats.Add(String.Format("{0} - {1}", thisNode, cat));
                }
                WriteObject(cats);

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            ListCategories(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Categories NOT returned. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }
        }

    }

    [Cmdlet(VerbsCommon.Get, "CSAttributeValues")]
    public class GetCSAttributeValuesCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 CategoryID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        SoapApi connection;
        Boolean recurse;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSAttributeValuesCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                ListAttributes(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Attributes NOT returned. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "GetCSAttributeValuesCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void ListAttributes(Int64 thisNode)
        {
            try
            {

                // copy the category and add the log entry
                Dictionary<String, List<Object>> response = connection.ListAttributes(thisNode, CategoryID);
                WriteObject(response);

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            ListAttributes(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Categories NOT returned. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSCategory")]
    public class AddCSCategoryCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 CategoryID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        SoapApi connection;
        Boolean recurse;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSCategoryCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                AddCategory(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Category {1} NOT added. ERROR: {2}", NodeID, CategoryID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSCategoryCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void AddCategory(Int64 thisNode)
        {
            try
            {

                // copy the category and add the log entry
                connection.AddCategoryToNode(thisNode, CategoryID, false, false, false);
                WriteObject(String.Format("{0} - Category {1} added", thisNode, CategoryID));

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            AddCategory(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - Category {1} NOT added. ERROR: {2}", thisNode, CategoryID, e.Message);
                WriteObject(message);
            }
        }

    }

    #endregion

    #region Member

    [Cmdlet(VerbsCommon.Add, "CSUser")]
    public class AddCSUserCommand : Cmdlet
    {

        #region Parameters and globals

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

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                String message = String.Format("{0} - user NOT created. ERROR: {1}", Login, e.Message);
                WriteObject(message);
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

    [Cmdlet(VerbsCommon.Add, "CSGroup")]
    public class AddCSGroupCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 LeaderID { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSGroupCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                // create the group
                Int64 response = connection.CreateGroup(Name, LeaderID);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - group NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
                ErrorRecord err = new ErrorRecord(e, "AddCSGroupCommand", ErrorCategory.NotSpecified, this);
                WriteError(err);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSGroupCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Remove, "CSMember")]
    public class RemoveCSMemberCommand : Cmdlet 
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 MemberID { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

                // create the user
                String message = String.Format("{0} - {1}", MemberID, connection.DeleteMember(MemberID));

                // write the output
                WriteObject(message);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - User NOT deleted. ERROR: {1}", MemberID, e.Message);
                WriteObject(message);
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

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Login { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                String message = String.Format("{0} - ID NOT retrieved. ERROR: {1}", Login, e.Message);
                WriteObject(message);
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

    /*
    [Cmdlet(VerbsCommon.Add, "CSMemberToGroup")]
    public class AddCSMemberToGroup : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 GroupID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 MemberID { get; set; }

        SoapAPI connection;

        #endregion

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
                connection = new SoapAPI();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSMemberToGroup", ErrorCategory.NotSpecified, this);
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
                connection = new SoapAPI();

                // create the user
                String message = String.Format("{0} - {2} member {1} to group", GroupID, MemberID, connection.AddMemberToGroup(GroupID, MemberID));

                // write the output
                WriteObject(message);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - member {1} NOT added to group. ERROR: {2}", GroupID, MemberID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSMemberToGroup", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }
    */

    #endregion

    #region Classifications

    [Cmdlet(VerbsCommon.Add, "CSClassifications")]
    public class AddCSClassificationsCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64[] ClassificationIDs { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                AddClassifications(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - classifications NOT applied. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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

        internal void AddClassifications(Int64 thisNode)
        {

            try
            {

                // add the classifications and return the message
                String message = "";
                Boolean success = connection.AddClassifications(thisNode, ClassificationIDs);
                if (success)
                    message = String.Format("{0} - classifications applied", thisNode);
                else
                    message = String.Format("{0} - classifications NOT applied. ERROR: unknown", thisNode);
                WriteObject(message);

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            AddClassifications(child);
                        }
                    }
                }

            }
            catch (Exception e)
            {
                String message = String.Format("{0} - classifications NOT applied. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    #endregion

    #region Records management

    [Cmdlet(VerbsCommon.Add, "CSRMClassification")]
    public class AddCSRMClassificationCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 RMClassificationID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                AddRMClassification(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - RM classification NOT applied. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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

        internal void AddRMClassification(Int64 thisNode)
        {

            try
            {

                // add the RM classification
                String message = "";
                Boolean success = connection.AddRMClassification(thisNode, RMClassificationID);
                if (success)
                {
                    message = String.Format("{0} - RM classification applied", thisNode);
                }
                else
                {
                    message = String.Format("{0} - RM classification NOT applied. ERROR: unknown.", thisNode);
                }

                // write the output
                WriteObject(message);

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            AddRMClassification(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - RM classification NOT applied. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    [Cmdlet(VerbsCommon.Set, "CSFinaliseRecord")]
    public class SetCSFinaliseRecordCommand : Cmdlet
    {

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

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
                FinaliseRecord(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT finalised. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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

        internal void FinaliseRecord(Int64 thisNode)
        {

            try
            {

                // finalise the record and return the message
                String message = "";
                Boolean success = connection.FinaliseRecord(thisNode);
                if (success)
                    message = String.Format("{0} - item finalised", thisNode);
                else
                    message = String.Format("{0} - item NOT finalised. ERROR: unknown", thisNode);
                WriteObject(message);

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            FinaliseRecord(child);
                        }
                    }
                }

            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT finalised. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    #endregion

    #region Physical objects

    [Cmdlet(VerbsCommon.Add, "CSPhysItem")]
    public class AddCSPhysItemCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 PhysicalItemSubType { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String HomeLocation { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Description { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String UniqueID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Keywords { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LocatorType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String RefRate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String OffsiteStorageID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String ClientName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String TemporaryID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LabelType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 ClientID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfCopies { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfLabels { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfItems { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean GenerateLabel { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime FromDate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime ToDate { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysItemCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                Int64 response = connection.CreatePhysicalItem(Name, ParentID, Globals.PhysicalItemTypes.PhysicalItem, PhysicalItemSubType, HomeLocation, Description, UniqueID, Keywords, LocatorType,
                    RefRate, OffsiteStorageID, ClientName, TemporaryID, LabelType, ClientID, NumberOfCopies, NumberOfLabels, NumberOfItems, GenerateLabel, FromDate, ToDate);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysItemCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSPhysContainer")]
    public class AddCSPhysContainerCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 PhysicalItemSubType { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String HomeLocation { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Description { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String UniqueID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Keywords { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LocatorType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String RefRate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String OffsiteStorageID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String ClientName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String TemporaryID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LabelType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 ClientID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfCopies { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfLabels { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfItems { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean GenerateLabel { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime FromDate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime ToDate { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysContainerCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the item and write the response
                Int64 response = connection.CreatePhysicalItem(Name, ParentID, Globals.PhysicalItemTypes.PhysicalItemContainer, PhysicalItemSubType, HomeLocation, Description, UniqueID, Keywords, LocatorType,
                    RefRate, OffsiteStorageID, ClientName, TemporaryID, LabelType, ClientID, NumberOfCopies, NumberOfLabels, NumberOfItems, GenerateLabel, FromDate, ToDate);
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysContainerCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSPhysBox")]
    public class AddCSPhysBoxCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 PhysicalItemSubType { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String HomeLocation { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Description { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String UniqueID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Keywords { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LocatorType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String RefRate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String OffsiteStorageID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String ClientName { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String TemporaryID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String LabelType { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 ClientID { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfCopies { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfLabels { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Int64 NumberOfItems { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public Boolean GenerateLabel { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime FromDate { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public DateTime ToDate { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysBoxCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the item and write the response
                Int64 response = connection.CreatePhysicalItem(Name, ParentID, Globals.PhysicalItemTypes.PhysicalItemBox, PhysicalItemSubType, HomeLocation, Description, UniqueID, Keywords, LocatorType,
                    RefRate, OffsiteStorageID, ClientName, TemporaryID, LabelType, ClientID, NumberOfCopies, NumberOfLabels, NumberOfItems, GenerateLabel, FromDate, ToDate);
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - item NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSPhysBoxCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Set, "CSPhysObjToBox")]
    public class SetCSPhysObjToBoxCommand : Cmdlet
    {

        #region Parameters and globals

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

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSPhysObjToBoxCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                String message = "";

                // convert any null booleans - defaults to false
                Boolean blnUpdateLocation = UpdateLocation ?? false;
                Boolean blnUpdateRSI = UpdateRSI ?? false;
                Boolean blnUpdateStatus = UpdateStatus ?? false;

                // assign to the box
                Boolean success = connection.AssignToBox(ItemID, BoxID, blnUpdateLocation, blnUpdateRSI, blnUpdateStatus);
                if (success)
                {
                    message = String.Format("{0} - assigned to box {1}", ItemID, BoxID);
                }
                else
                {
                    message = String.Format("{0} - NOT assigned to box {1}. ERROR: unknown", ItemID, BoxID);
                }

                // write the output
                WriteObject(message);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - NOT assigned to box {1}. ERROR: {2}", ItemID, BoxID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "SetCSPhysObjToBoxCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion

    #region Permissions

    [Cmdlet(VerbsCommon.Get, "CSPermissions")]
    public class GetCSPermissions : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String[] Role { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public Int64? UserID { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "GetCSPermissions", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                Int64 userID = UserID ?? 0;
                String[] responses = connection.GetPermissions(NodeID, Role, userID);
                foreach (String response in responses)
                {
                    WriteObject(response);
                }
            }
            catch (Exception e)
            {
                WriteObject(String.Format("{0} - permissions NOT retrieved. ERROR: {1}", NodeID, e.Message));
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
                ErrorRecord err = new ErrorRecord(e, "GetCSPermissions", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    [Cmdlet(VerbsCommon.Set, "CSOwner")]
    public class SetCSOwner : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public Int64 UserID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public String[] Permissions { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSOwner", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                UpdateOwner(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - owner NOT updated. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "SetCSOwner", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void UpdateOwner(Int64 thisNode)
        {

            try
            {

                // add the assigned access
                try
                {
                    connection.SetOwner(thisNode, UserID, Permissions);
                    WriteObject(String.Format("{0} - owner updated", thisNode));
                }
                catch (Exception e)
                {
                    WriteObject(String.Format("{0} - owner NOT updated. ERROR: {1}", thisNode, e.Message));
                }

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            UpdateOwner(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - owner NOT updated. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    [Cmdlet(VerbsCommon.Set, "CSOwnerGroup")]
    public class SetCSOwnerGroup : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public Int64 GroupID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public String[] Permissions { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSOwnerGroup", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                UpdateOwnerGroup(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - owner group NOT updated. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "SetCSOwnerGroup", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void UpdateOwnerGroup(Int64 thisNode)
        {

            try
            {

                // add the assigned access
                try
                {
                    connection.SetOwnerGroup(thisNode, GroupID, Permissions);
                    WriteObject(String.Format("{0} - owner group updated", thisNode));
                }
                catch (Exception e)
                {
                    WriteObject(String.Format("{0} - owner group NOT updated. ERROR: {1}", thisNode, e.Message));
                }

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            UpdateOwnerGroup(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - owner group NOT updated. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    [Cmdlet(VerbsCommon.Set, "CSPublicAccess")]
    public class SetCSPublicAccess : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String[] Permissions { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSPublicAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                UpdatePublicAccess(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - public access NOT updated. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "SetCSPublicAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void UpdatePublicAccess(Int64 thisNode)
        {

            try
            {

                // add the assigned access
                try
                {
                    connection.SetPublicAccess(thisNode, Permissions);
                    WriteObject(String.Format("{0} - public access updated", thisNode));
                }
                catch (Exception e)
                {
                    WriteObject(String.Format("{0} - public access NOT updated. ERROR: {1}", thisNode, e.Message));
                }

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            UpdatePublicAccess(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - public access NOT updated. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    [Cmdlet(VerbsCommon.Set, "CSAssignedAccess")]
    public class SetCSAssignedAccess: Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 UserID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String[] Permissions { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "SetCSAssignedAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                AddAssignedAccess(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - permissions NOT assigned. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "SetCSAssignedAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void AddAssignedAccess(Int64 thisNode)
        {

            try
            {

                // add the assigned access
                try
                {
                    connection.SetAssignedAccess(thisNode, UserID, Permissions);
                    WriteObject(String.Format("{0} - permissions applied", thisNode));
                }
                catch (Exception e)
                {
                    WriteObject(String.Format("{0} - permissions NOT applied. ERROR: {1}", thisNode, e.Message));
                }

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            AddAssignedAccess(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - permissions NOT applied. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }
        
    }

    [Cmdlet(VerbsCommon.Remove, "CSAssignedAccess")]
    public class RemoveCSAssignedAccess : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 UserID { get; set; }
        [Parameter(ValueFromPipelineByPropertyName = true)]
        public SwitchParameter Recurse
        {
            get { return recurse; }
            set { recurse = value; }
        }

        private Boolean recurse;
        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "RemoveCSAssignedAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                RemoveAssignedAccess(NodeID);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - permissions NOT removed. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "RemoveCSAssignedAccess", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        internal void RemoveAssignedAccess(Int64 thisNode)
        {

            try
            {

                // add the assigned access
                try
                {
                    connection.RemoveAssignedAccess(thisNode, UserID);
                    WriteObject(String.Format("{0} - permissions removed", thisNode));
                }
                catch (Exception e)
                {
                    WriteObject(String.Format("{0} - permissions NOT removed. ERROR: {1}", thisNode, e.Message));
                }

                // are we recursing through this object?
                if (Recurse)
                {
                    List<Int64> children = connection.GetChildren(thisNode);
                    if (children.Count > 0)
                    {
                        foreach (Int64 child in children)
                        {
                            RemoveAssignedAccess(child);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - permissions NOT applied. ERROR: {1}", thisNode, e.Message);
                WriteObject(message);
            }

        }

    }

    #endregion

    #region Template Workspaces

    [Cmdlet(VerbsCommon.Add, "CSBinder")]
    public class AddCSBinderCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 TemplateID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ClassificationID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Description { get; set; }

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.RestConnectionOpened)
                {
                    ThrowTerminatingError(Errors.RestConnectionMissing(this));
                    return;
                }

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSBinderCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                Int64 response = RestApi.CreateFromTemplate(TemplateID, ParentID, ClassificationID, Name, Description);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - binder NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

    }

    [Cmdlet(VerbsCommon.Add, "CSCase")]
    public class AddCSCaseCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 TemplateID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ParentID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 ClassificationID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Name { get; set; }
        [Parameter(Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public String Description { get; set; }

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.RestConnectionOpened)
                {
                    ThrowTerminatingError(Errors.RestConnectionMissing(this));
                    return;
                }

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSCaseCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {

                Int64 response = RestApi.CreateFromTemplate(TemplateID, ParentID, ClassificationID, Name, Description);

                // write the output
                WriteObject(response);
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - case NOT created. ERROR: {1}", Name, e.Message);
                WriteObject(message);
            }

        }

        protected override void EndProcessing()
        {
            base.EndProcessing();
        }

    }

    #endregion
 
    #region Renditions

    [Cmdlet(VerbsCommon.Add, "CSRendition")]
    public class AddCSRenditionCommand : Cmdlet
    {

        #region Parameters and globals

        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 NodeID { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public Int64 Version { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Type { get; set; }
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty]
        public String Document { get; set; }

        SoapApi connection;

        #endregion

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            try
            {
                // create the connection object
                if (!Globals.SoapConnectionOpened)
                {
                    ThrowTerminatingError(Errors.SoapConnectionMissing(this));
                    return;
                }
                connection = new SoapApi();

            }
            catch (Exception e)
            {
                ErrorRecord err = new ErrorRecord(e, "AddCSRenditionCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            try
            {
                // create the folder
                Int64 response = connection.CreateRendition(NodeID, Version, Type, Document);

                // write the output
                WriteObject("Rendition added");
            }
            catch (Exception e)
            {
                String message = String.Format("{0} - rendition NOT added. ERROR: {1}", NodeID, e.Message);
                WriteObject(message);
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
                ErrorRecord err = new ErrorRecord(e, "AddCSRenditionCommand", ErrorCategory.NotSpecified, this);
                ThrowTerminatingError(err);
            }
        }

    }

    #endregion
    
    internal class Errors
    {

        internal static ErrorRecord SoapConnectionMissing(Object Object)
        {
            String msg = "Connection has not been opened. Please open the connection first using 'Open-CSConnection'";
            Exception exception = new Exception(msg);
            ErrorRecord err = new ErrorRecord(exception, "ConnectionMissing", ErrorCategory.ResourceUnavailable, Object);
            return err;
        }

        internal static ErrorRecord RestConnectionMissing(Object Object)
        {
            String msg = "Connection has not been opened. Please open the connection first using 'Open-CSConnectionRest'";
            Exception exception = new Exception(msg);
            ErrorRecord err = new ErrorRecord(exception, "ConnectionMissing", ErrorCategory.ResourceUnavailable, Object);
            return err;
        }

    }

}
