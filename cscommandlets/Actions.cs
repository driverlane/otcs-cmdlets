using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;

namespace CGIPSTools
{

    internal class Globals
    {
        internal static String Username;
        internal static String Password;
        internal static String ServicesDirectory;
        internal static Boolean Opened;
    }

    internal class Connection
    {

        #region Class housekeeping

        internal String ErrorMessage;

        private String AuthenticationEndpointAddress = "Authentication.svc";
        private String CollaborationEndpointAddress = "Collaboration.svc";
        private String DocumentManagementEndpointAddress = "DocumentManagement.svc";

        private Authentication.AuthenticationClient authClient;
        private Collaboration.CollaborationClient collabClient;
        private Collaboration.OTAuthentication collabAuth;
        private DocumentManagement.DocumentManagementClient docClient;
        private DocumentManagement.OTAuthentication docAuth;

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

            InitialiseClients();
        }

        internal void InitialiseClients()
        {
            ErrorMessage = null;

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

                Globals.Opened = true;
            }
            catch (Exception e)
            {
                ErrorMessage = "Connection error: " + e.Message;
                Globals.Opened = false;
            }

        }

        #endregion

        #region Internal methods

        internal AGANode CreateContainer(String Name, Int32 ParentID, ObjectType ObjectType)
        {
            AGANode newNode = new AGANode();
            ErrorMessage = null;

            switch (ObjectType)
            {
                case ObjectType.Folder:
                    try
                    {
                        DocumentManagement.Metadata metadata = new DocumentManagement.Metadata();
                        DocumentManagement.Node node = docClient.CreateFolder(ref docAuth, ParentID, Name, "", metadata);
                        newNode.ID = node.ID;
                        // AGA web part property in the format -2000:94485::Enterprise » Testing Sandpit » atestmt
                        newNode.NodeValue = String.Format("{0}:{1}::{2}", node.VolumeID, node.ID, GetNodePath(newNode.ID));
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                DocumentManagement.Node node = docClient.GetNodeByName(ref docAuth, ParentID, Name);
                                newNode.ID = node.ID;
                                // AGA web part property in the format -2000:94485::Enterprise » Testing Sandpit » atestmt
                                newNode.NodeValue = String.Format("{0}:{1}::{2}", node.VolumeID, node.ID, GetNodePath(newNode.ID));
                                newNode.Message = e.Message + " Returning existing node details.";
                            }
                            catch (Exception e2)
                            {
                                ErrorMessage = e2.Message;
                            }
                        }
                        else
                        {
                            ErrorMessage = e.Message;
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
                        newNode.ID = node.ID;
                        // AGA web part property in the format -2000:94485::Enterprise » Testing Sandpit » atestmt
                        newNode.NodeValue = String.Format("{0}:{1}::{2}", node.VolumeID, node.ID, GetNodePath(newNode.ID));
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                DocumentManagement.Node node = docClient.GetNodeByName(ref docAuth, ParentID, Name);
                                newNode.ID = node.ID;
                                // AGA web part property in the format -2000:94485::Enterprise » Testing Sandpit » atestmt
                                newNode.NodeValue = String.Format("{0}:{1}::{2}", node.VolumeID, node.ID, GetNodePath(newNode.ID));
                                newNode.Message = e.Message + " Returning existing node details.";
                            }
                            catch (Exception e2)
                            {
                                ErrorMessage = e2.Message;
                            }
                        }
                        else
                        {
                            ErrorMessage = e.Message;
                        }
                    }
                    break;

            }

            return newNode;
        }

        internal void UpdateProjectFromTemplate(Int32 ProjectID, Int32 TemplateID)
        {
            CopyProjectParticipants(ProjectID, TemplateID);
            CopyChildren(ProjectID, TemplateID, ObjectType.Project);
        }

        #endregion

        #region Private methods

        private void CopyProjectParticipants(Int32 ProjectID, Int32 TemplateID)
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
                ErrorMessage = e.Message;
            }
        }

        private void CopyChildren(Int32 Copy, Int32 Template, ObjectType ObjectType)
        {
            ErrorMessage = null;
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
                ErrorMessage = e.Message;
            }

        }

        private String GetNodePath(Int32 NodeID)
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

        private String GetParentName(Int32 NodeID)
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
