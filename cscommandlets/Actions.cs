using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;

namespace cscommandlets
{

    internal class Globals
    {
        internal static String Username;
        internal static String Password;
        internal static String ServicesDirectory;
        internal static Boolean ConnectionOpened;
    }

    internal class Connection
    {

        #region Class housekeeping

        private Boolean LargeIDs = true;

        private String AuthenticationEndpointAddress = "Authentication.svc";
        private String CollaborationEndpointAddress = "Collaboration.svc";
        private String DocumentManagementEndpointAddress = "DocumentManagement.svc";
        private String MemberServiceEndpointAddress = "MemberService.svc";
        private String ClassificationsEndpointAddress = "Classifications.svc";

        private Authentication.AuthenticationClient authClient;
        private Collaboration.CollaborationClient collabClient;
        private Collaboration.OTAuthentication collabAuth;
        private DocumentManagement.DocumentManagementClient docClient;
        private DocumentManagement.OTAuthentication docAuth;
        private MemberService.MemberServiceClient memberClient;
        private MemberService.OTAuthentication memberAuth;
        private Classifications.ClassificationsClient classClient;
        private Classifications.OTAuthentication classAuth;

        internal enum ObjectType { Folder, Project };

        internal Connection()
        {
            if (!Globals.ServicesDirectory.EndsWith("/"))
            {
                Globals.ServicesDirectory = Globals.ServicesDirectory + "/";
            }
            AuthenticationEndpointAddress = Globals.ServicesDirectory + AuthenticationEndpointAddress;
            CollaborationEndpointAddress = Globals.ServicesDirectory + CollaborationEndpointAddress;
            DocumentManagementEndpointAddress = Globals.ServicesDirectory + DocumentManagementEndpointAddress;
            MemberServiceEndpointAddress = Globals.ServicesDirectory + MemberServiceEndpointAddress;
            ClassificationsEndpointAddress = Globals.ServicesDirectory + ClassificationsEndpointAddress;

            InitialiseClients();
        }

        internal void InitialiseClients()
        {

            // initialise the web service clients
            try
            {
                EndpointAddress authAddress = new EndpointAddress(AuthenticationEndpointAddress);
                BasicHttpBinding authBinding = new BasicHttpBinding();
                authBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                authBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                authClient = new Authentication.AuthenticationClient(authBinding, authAddress);

                EndpointAddress collabAddress = new EndpointAddress(CollaborationEndpointAddress);
                BasicHttpBinding collabBinding = new BasicHttpBinding();
                collabBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                collabBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                collabClient = new Collaboration.CollaborationClient(collabBinding, collabAddress);

                EndpointAddress docAddress = new EndpointAddress(DocumentManagementEndpointAddress);
                BasicHttpBinding docBinding = new BasicHttpBinding();
                docBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                docBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                docClient = new DocumentManagement.DocumentManagementClient(docBinding, docAddress);

                EndpointAddress memberAddress = new EndpointAddress(MemberServiceEndpointAddress);
                BasicHttpBinding memberBinding = new BasicHttpBinding();
                memberBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                memberBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                memberClient = new MemberService.MemberServiceClient(memberBinding, memberAddress);

                EndpointAddress classAddress = new EndpointAddress(ClassificationsEndpointAddress);
                BasicHttpBinding classBinding = new BasicHttpBinding();
                classBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                classBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                classClient = new Classifications.ClassificationsClient(classBinding, classAddress);

                // get the authentication token and create the authentication object
                String password = Globals.Password;
                if (password.StartsWith("!=!enc!=!"))
                {
                    password = Encryption.DecryptString(password);
                }
                String token = authClient.AuthenticateUser(Globals.Username, password);
                collabAuth = new Collaboration.OTAuthentication();
                collabAuth.AuthenticationToken = token;
                docAuth = new DocumentManagement.OTAuthentication();
                docAuth.AuthenticationToken = token;
                memberAuth = new MemberService.OTAuthentication();
                memberAuth.AuthenticationToken = token;
                classAuth = new Classifications.OTAuthentication();
                classAuth.AuthenticationToken = token;

                Globals.ConnectionOpened = true;
            }
            catch (Exception e)
            {
                Globals.ConnectionOpened = false;
                throw e;
            }

        }

        #endregion

        #region Internal methods

        internal Int64 CreateContainer(String Name, Int64 ParentID, ObjectType ObjectType)
        {
            Int64 newNode = 0;

            switch (ObjectType)
            {
                case ObjectType.Folder:
                    try
                    {
                        DocumentManagement.Metadata metadata = new DocumentManagement.Metadata();
                        DocumentManagement.Node node = docClient.CreateFolder(ref docAuth, ParentID, Name, "", metadata);
                        newNode = node.ID;
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                DocumentManagement.Node node = docClient.GetNodeByName(ref docAuth, ParentID, Name);
                                newNode = node.ID;
                                // todo set up a non terminating error
                            }
                            catch (Exception e2)
                            {
                                throw e2;
                            }
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    break;

                case ObjectType.Project:
                    try
                    {
                        Collaboration.ProjectInfo projectInfo = new Collaboration.ProjectInfo();
                        projectInfo.Name = Name;
                        projectInfo.ParentID = ParentID;
                        Collaboration.Node node = collabClient.CreateProject(ref collabAuth, projectInfo);
                        newNode = node.ID;
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                DocumentManagement.Node node = docClient.GetNodeByName(ref docAuth, ParentID, Name);
                                newNode = node.ID;
                                // todo set up a non terminating error
                            }
                            catch (Exception e2)
                            {
                                throw e2;
                            }
                        }
                        else
                        {
                            throw e;
                        }
                    }
                    break;

            }

            return newNode;
        }

        internal void UpdateProjectFromTemplate(Int64 ProjectID, Int64 TemplateID)
        {
            CopyProjectParticipants(ProjectID, TemplateID);
            CopyChildren(ProjectID, TemplateID, ObjectType.Project);
        }

        internal String DeleteNode(Int64 NodeID)
        {
            try
            {
                docClient.DeleteNode(ref docAuth, NodeID);
                return "Deleted";
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        internal Int64 CreateUser(String Login, Int64 DepartmentGroupID, String Password, String FirstName, String MiddleName, String LastName, String Email, String Fax, String OfficeLocation,
            String Phone, String Title, Boolean LoginEnabled, Boolean PublicAccessEnabled, Boolean CreateUpdateUsers, Boolean CreateUpdateGroups, Boolean CanAdministerUsers, Boolean CanAdministerSystem)
        {
            try
            {
                MemberService.User user = new MemberService.User();
                user.Name = Login;
                user.DepartmentGroupID = DepartmentGroupID;
                user.Password = Password;
                user.FirstName = FirstName;
                user.MiddleName = MiddleName;
                user.LastName = LastName;
                user.Email = Email;
                user.Fax = Fax;
                user.OfficeLocation = OfficeLocation;
                user.Phone = Phone;
                //user.TimeZone = TZone;
                user.Title = Title;

                // set up privileges
                if (user.Privileges == null) user.Privileges = new MemberService.MemberPrivileges();
                user.Privileges.LoginEnabled = LoginEnabled;
                user.Privileges.PublicAccessEnabled = PublicAccessEnabled;
                user.Privileges.CreateUpdateUsers = CreateUpdateUsers;
                user.Privileges.CreateUpdateGroups = CreateUpdateGroups;
                user.Privileges.CanAdministerUsers = CanAdministerUsers;
                user.Privileges.CanAdministerSystem = CanAdministerSystem;

                Int64 response = memberClient.CreateUser(ref memberAuth, user);
                return response;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        internal String DeleteUser(Int64 UserID)
        {
            try
            {
                memberClient.DeleteMember(ref memberAuth, UserID);
                return "Deleted";
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        internal Boolean AddClassifications(Int32 NodeID, Int32[] ClassIDs)
        {
            try
            {
                return classClient.ApplyClassifications(ref classAuth, NodeID, ClassIDs);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        #region Private methods

        private void CopyProjectParticipants(Int64 ProjectID, Int64 TemplateID)
        {

            try
            {
                // copy the public access
                Collaboration.ProjectInfo templateInfo = collabClient.GetProject(ref collabAuth, TemplateID);
                Collaboration.ProjectInfo projectInfo = collabClient.GetProject(ref collabAuth, ProjectID);
                projectInfo.PublicAccess = templateInfo.PublicAccess;

                // add the participants from the template
                Collaboration.ProjectParticipants templateParticipants = collabClient.GetParticipants(ref collabAuth, TemplateID);
                List<Collaboration.ProjectRoleUpdateInfo> participants = new List<Collaboration.ProjectRoleUpdateInfo>();

                if (templateParticipants.Coordinators != null)
                {
                    foreach (Collaboration.Member member in templateParticipants.Coordinators)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                            participant.RoleAction = Collaboration.ProjectRoleAction.ASSIGN;
                            participant.UserID = member.ID;
                            participant.Role = Collaboration.ProjectRole.COORDINATOR;
                            participants.Add(participant);
                        }
                    }
                }
                if (templateParticipants.Members != null)
                {
                    foreach (Collaboration.Member member in templateParticipants.Members)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                            participant.RoleAction = Collaboration.ProjectRoleAction.ASSIGN;
                            participant.UserID = member.ID;
                            participant.Role = Collaboration.ProjectRole.MEMBER;
                            participants.Add(participant);
                        }
                    }
                }
                if (templateParticipants.Guests != null)
                {
                    foreach (Collaboration.Member member in templateParticipants.Guests)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                            participant.RoleAction = Collaboration.ProjectRoleAction.ASSIGN;
                            participant.UserID = member.ID;
                            participant.Role = Collaboration.ProjectRole.GUEST;
                            participants.Add(participant);
                        }
                    }
                }

                // remove any inherited participants
                Collaboration.ProjectParticipants projectParticipants = collabClient.GetParticipants(ref collabAuth, ProjectID);
                if (projectParticipants.Coordinators != null)
                {
                    foreach (Collaboration.Member member in projectParticipants.Coordinators)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            if (!participants.Any(item => item.UserID == member.ID))
                            {
                                Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                                participant.RoleAction = Collaboration.ProjectRoleAction.REMOVE;
                                participant.UserID = member.ID;
                                participant.Role = Collaboration.ProjectRole.COORDINATOR;
                                participants.Add(participant);
                            }
                        }
                    }
                }
                if (projectParticipants.Members != null)
                {
                    foreach (Collaboration.Member member in projectParticipants.Members)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            if (!participants.Any(item => item.UserID == member.ID))
                            {
                                Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                                participant.RoleAction = Collaboration.ProjectRoleAction.REMOVE;
                                participant.UserID = member.ID;
                                participant.Role = Collaboration.ProjectRole.MEMBER;
                                participants.Add(participant);
                            }
                        }
                    }
                }
                if (projectParticipants.Guests != null)
                {
                    foreach (Collaboration.Member member in projectParticipants.Guests)
                    {
                        if (member.Type != "ProjectGroup")
                        {
                            if (!participants.Any(item => item.UserID == member.ID))
                            {
                                Collaboration.ProjectRoleUpdateInfo participant = new Collaboration.ProjectRoleUpdateInfo();
                                participant.RoleAction = Collaboration.ProjectRoleAction.REMOVE;
                                participant.UserID = member.ID;
                                participant.Role = Collaboration.ProjectRole.GUEST;
                                participants.Add(participant);
                            }
                        }
                    }
                }
                collabClient.UpdateProjectParticipants(ref collabAuth, ProjectID, participants.ToArray());
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void CopyChildren(Int64 Copy, Int64 Template, ObjectType ObjectType)
        {
            try
            {
                DocumentManagement.Node parentNode = docClient.GetNode(ref docAuth, Copy);
                DocumentManagement.Node newParentNode = docClient.GetNode(ref docAuth, Template);

                if (parentNode.IsContainer && newParentNode.IsContainer)
                {
                    // change the ID if it's a volume object
                    Boolean VolumeObject = false;
                    if (ObjectType.Equals(ObjectType.Project)) { VolumeObject = true; }
                    if (VolumeObject)
                    {
                        Copy = -Copy;
                        Template = -Template;
                    }

                    DocumentManagement.Node[] children = docClient.ListNodes(ref docAuth, Template, true);
                    if (children != null)
                    {
                        foreach (DocumentManagement.Node child in children)
                        {
                            DocumentManagement.CopyOptions copyOptions = new DocumentManagement.CopyOptions();
                            docClient.CopyNode(ref docAuth, child.ID, Copy, child.Name, copyOptions);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        private String GetNodePath(Int64 NodeID)
        {

            String[] rootNodes = docClient.GetRootNodeTypes(ref docAuth);
            DocumentManagement.Node child = docClient.GetNode(ref docAuth, NodeID);
            DocumentManagement.Node parent = docClient.GetNode(ref docAuth, child.ParentID);
            String path = child.Name;
            path = parent.Name + " >> " + path;

            do
            {
                child = docClient.GetNode(ref docAuth, parent.ID);
                parent = docClient.GetNode(ref docAuth, child.ParentID);

                path = parent.Name + " >> " + path;

            } while (!rootNodes.Contains(parent.Name + "WS"));

            return path;
        }

        private String GetParentName(Int64 NodeID)
        {
            String name;
            DocumentManagement.Node child = docClient.GetNode(ref docAuth, NodeID);
            DocumentManagement.Node parent = docClient.GetNode(ref docAuth, child.ParentID);
            name = parent.Name;
            return name;
        }

        #endregion

    }

}
