using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using cscmdlets.DocumentManagement;

namespace cscmdlets
{

    internal class Globals
    {
        internal static String Username;
        internal static String Password;
        internal static String ServicesDirectory;
        internal static Boolean ConnectionOpened;

        internal enum PhysicalItemTypes
        {
            PhysicalItem,
            PhysicalItemContainer,
            PhysicalItemBox
        }

    }

    internal class Server
    {

        /* There's minimal error capture in this class.  Handle the error in the calling method unless it's necessary at this level (and it sometimes is). */

        #region API connect

        internal List<Exception> NonTerminatingExceptions;

        private String AuthenticationEndpointAddress = "Authentication.svc";
        private String CollaborationEndpointAddress = "Collaboration.svc";
        private String DocumentManagementEndpointAddress = "DocumentManagement.svc";
        private String MemberServiceEndpointAddress = "MemberService.svc";
        private String ClassificationsEndpointAddress = "Classifications.svc";
        private String RecordsManagementEndpointAddress = "RecordsManagement.svc";
        private String PhysicalObjectsEndpointAddress = "PhysicalObjects.svc";

        private Authentication.AuthenticationClient authClient;
        private Collaboration.CollaborationClient collabClient;
        private Collaboration.OTAuthentication collabAuth;
        private DocumentManagement.DocumentManagementClient docClient;
        private DocumentManagement.OTAuthentication docAuth;
        private MemberService.MemberServiceClient memberClient;
        private MemberService.OTAuthentication memberAuth;
        private Classifications.ClassificationsClient classClient;
        private Classifications.OTAuthentication classAuth;
        private RecordsManagement.RecordsManagementClient recmanClient;
        private RecordsManagement.OTAuthentication recmanAuth;
        private PhysicalObjects.PhysicalObjectsClient poClient;
        private PhysicalObjects.OTAuthentication poAuth;

        internal enum ObjectType { Folder, Project };

        internal Server()
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
            RecordsManagementEndpointAddress = Globals.ServicesDirectory + RecordsManagementEndpointAddress;
            PhysicalObjectsEndpointAddress = Globals.ServicesDirectory + PhysicalObjectsEndpointAddress;

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

                EndpointAddress recmanAddress = new EndpointAddress(RecordsManagementEndpointAddress);
                BasicHttpBinding recmanBinding = new BasicHttpBinding();
                recmanBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                recmanBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                recmanClient = new RecordsManagement.RecordsManagementClient(recmanBinding, recmanAddress);

                EndpointAddress poAddress = new EndpointAddress(PhysicalObjectsEndpointAddress);
                BasicHttpBinding poBinding = new BasicHttpBinding();
                poBinding.SendTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                poBinding.OpenTimeout = new TimeSpan(0, 0, 0, 0, 100000);
                poClient = new PhysicalObjects.PhysicalObjectsClient(poBinding, poAddress);

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
                recmanAuth = new RecordsManagement.OTAuthentication();
                recmanAuth.AuthenticationToken = token;
                poAuth = new PhysicalObjects.OTAuthentication();
                poAuth.AuthenticationToken = token;

                Globals.ConnectionOpened = true;
            }
            catch (Exception e)
            {
                Globals.ConnectionOpened = false;
                throw e;
            }

        }

        internal void CloseClients()
        {
            collabClient.Close();
            docClient.Close();
            memberClient.Close();
            classClient.Close();
            recmanClient.Close();
        }

        #endregion

        #region API calls

        #region Docman calls

        internal Int64 CreateContainer(String Name, Int64 ParentID, ObjectType ObjectType)
        {
            Int64 newNode = 0;
            NonTerminatingExceptions = new List<Exception>();

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

                                // log this error, not throw it
                                NonTerminatingExceptions.Add(e);
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

                                // log this error, not throw it
                                NonTerminatingExceptions = new List<Exception>();
                                NonTerminatingExceptions.Add(e);
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

        internal String DeleteNode(Int64 NodeID)
        {
            docClient.DeleteNode(ref docAuth, NodeID);
            return "Deleted";
        }

        #region cats and atts

        /*
        internal void UpdateAttribute(Int64 NodeID, Int64 CategoryID, String Attribute, Object[] Values, Object[] Replace, Boolean AddAttribute)
        {

            // get the current metadata
            Node item = docClient.GetNode(ref docAuth, NodeID);
            Metadata metadata = item.Metadata;

            // create the attribute
            DataValue attribute = CreateDataValue(CategoryID, Attribute, Values);

            // clean the metadata and add the new attribute
            item.Metadata = CleanMetadata(metadata, attribute, Replace, "", AddAttribute);

            // update the object
            docClient.UpdateNode(ref docAuth, item);

        }

        internal void ClearAttribute(Int64 NodeID, Int64 CategoryID, String Attribute)
        {

            // get the current metadata
            Node item = docClient.GetNode(ref docAuth, NodeID);
            Metadata metadata = item.Metadata;

            // get the ID of the attribute to clear
            String attID = GetAttributeID(CategoryID, Attribute);

            // clean the metadata and the attribute
            item.Metadata = CleanMetadata(metadata, null, null, attID, false);

            // update the object
            docClient.UpdateNode(ref docAuth, item);

        }

        internal void CopyCategories(Int64 SourceID, Int64 TargetID, Boolean MergeCats, Boolean MergeAtts)
        {

            // get the metadata from the source
            Node source = docClient.GetNode(ref docAuth, SourceID);

            // get the metadata from the target
            Node target = docClient.GetNode(ref docAuth, TargetID);

            // combine categories
            target.Metadata = CombineCats(source.Metadata, target.Metadata, MergeCats, MergeAtts);

            // update the object
            docClient.UpdateNode(ref docAuth, target);

        }

        internal void CopyCategory(Int64 SourceID, Int64 TargetID, Boolean MergeAtts)
        {

            // get the metadata from the source
            Node source = docClient.GetNode(ref docAuth, SourceID);

            // get the metadata from the target
            Node target = docClient.GetNode(ref docAuth, TargetID);

            // combine categories
            target.Metadata = CombineCat(source.Metadata, target.Metadata, MergeAtts);

            // update the object
            docClient.UpdateNode(ref docAuth, target);

        }

        private Metadata CombineCats(Metadata Source, Metadata Target, Boolean MergeCats, Boolean MergeAtts)
        {

        }

        private Metadata CombineCat(Metadata Source, Metadata Target, Boolean MergeAtts)
        {

        }

        private AttributeGroup CombineAttributeGroup(AttributeGroup Source, AttributeGroup Target, Boolean MergeAtts)
        {

        }

        private Metadata UpdateMetadata(Metadata Metadata, List<AttributeGroup> Categories, Dictionary<String, Object[]> Replace, Boolean Add)
        {

            // assume they're all cleaned

            // iterate the cats on metadata, process any that are common

                // replace/add the att, checking if there's a constraint on the replace

            // if we're adding, add the new cats

        }
        */

        #endregion

        #endregion

        #region Member calls

        internal Int64 CreateUser(String Login, Int64 DepartmentGroupID, String Password, String FirstName, String MiddleName, String LastName, String Email, String Fax, String OfficeLocation,
            String Phone, String Title, Boolean LoginEnabled, Boolean PublicAccessEnabled, Boolean CreateUpdateUsers, Boolean CreateUpdateGroups, Boolean CanAdministerUsers, Boolean CanAdministerSystem)
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

        internal String DeleteUser(Int64 UserID)
        {
            memberClient.DeleteMember(ref memberAuth, UserID);
            return "Deleted";
        }

        internal Int64 GetUserIDByLoginName(String Login)
        {
            MemberService.Member user = memberClient.GetMemberByLoginName(ref memberAuth, Login);
            return user.ID;
        }

        #endregion

        #region Classifications calls

        internal Boolean AddClassifications(Int64 NodeID, Int64[] ClassIDs)
        {
            return classClient.ApplyClassifications(ref classAuth, NodeID, ClassIDs);
        }

        #endregion

        #region RM calls

        internal Boolean AddRMClassification(Int64 NodeID, Int64 RMClassID)
        {
            RecordsManagement.RMAdditionalInfo info = new RecordsManagement.RMAdditionalInfo();
            Int64[] otherIDs = {};
            return recmanClient.RMApplyClassification(ref recmanAuth, NodeID, RMClassID, info, otherIDs);
        }

        internal Boolean FinaliseRecord(Int64 NodeID)
        {
            return recmanClient.RMDeclareRecord(ref recmanAuth, NodeID);
        }

        #endregion

        #region Physical objects calls

        internal Int64 CreatePhysicalItem(String Name, Int64 ParentID, Globals.PhysicalItemTypes Type, Int64 SubType, String HomeLocation, String Description, String UniqueID,
            String Keywords, String LocatorType, String RefRate, String OffsiteStorageID, String ClientName, String TemporaryID, String LabelType, Int64 ClientID, Int64 NumberOfCopies,
            Int64 NumberOfLabels, Int64 NumberOfItems, Boolean GenerateLabel, DateTime? FromDate, DateTime? ToDate)
        {

            // populate the standard fields
            DocumentManagement.Node newNode = new DocumentManagement.Node();
            newNode.Name = Name;
            newNode.ParentID = ParentID;
            newNode.Comment = Description;
            newNode.Type = Type.ToString();
            //newNode.DisplayType = Type.ToString();

            // create the metadata element and attribute groups
            DocumentManagement.AttributeGroup poAtts = CreatePhysicalObjectAttributeGroup(Type, SubType, HomeLocation, UniqueID, Keywords, LocatorType, RefRate, OffsiteStorageID, ClientName,
                TemporaryID, LabelType, ClientID, NumberOfCopies, NumberOfLabels, NumberOfItems, GenerateLabel, FromDate, ToDate);
            newNode.Metadata = new DocumentManagement.Metadata();
            newNode.Metadata.AttributeGroups = new DocumentManagement.AttributeGroup[] { poAtts };

            return docClient.CreateNode(ref docAuth, newNode).ID;
        }

        private DocumentManagement.AttributeGroup CreatePhysicalObjectAttributeGroup(Globals.PhysicalItemTypes Type, Int64 SubType, String HomeLocation, String UniqueID,
            String Keywords, String LocatorType, String RefRate, String OffsiteStorageID, String ClientName, String TemporaryID, String LabelType, Int64 ClientID, Int64 NumberOfCopies,
            Int64 NumberOfLabels, Int64 NumberOfItems, Boolean GenerateLabel, DateTime? FromDate, DateTime? ToDate)
        {

            // selectedMedia
            DocumentManagement.IntegerValue subType = new DocumentManagement.IntegerValue();
            subType.Key = "selectedMedia";
            subType.Values = new Int64?[] { SubType };

            // poLocation
            DocumentManagement.StringValue homeLocation = new DocumentManagement.StringValue();
            homeLocation.Key = "poLocation";
            homeLocation.Values = new String[] { HomeLocation };

            // poUniqueID
            DocumentManagement.StringValue uniqueID = new DocumentManagement.StringValue();
            uniqueID.Key = "poUniqueID";
            uniqueID.Values = new String[] { UniqueID };

            // poKeywords
            DocumentManagement.StringValue keywords = new DocumentManagement.StringValue();
            keywords.Key = "poKeywords";
            keywords.Values = new String[] { Keywords };

            // LocatorType
            DocumentManagement.StringValue locatorType = new DocumentManagement.StringValue();
            locatorType.Key = "LocatorType";
            locatorType.Values = new String[] { LocatorType };

            // RefRate
            DocumentManagement.StringValue refRate = new DocumentManagement.StringValue();
            refRate.Key = "RefRate";
            refRate.Values = new String[] { RefRate };

            // OffsiteStorID
            DocumentManagement.StringValue offsiteStorageID = new DocumentManagement.StringValue();
            offsiteStorageID.Key = "OffsiteStorID";
            offsiteStorageID.Values = new String[] { OffsiteStorageID };

            // Client_Name
            DocumentManagement.StringValue clientName = new DocumentManagement.StringValue();
            clientName.Key = "Client_Name";
            clientName.Values = new String[] { ClientName };

            // TemporaryID
            DocumentManagement.StringValue temporaryID = new DocumentManagement.StringValue();
            temporaryID.Key = "TemporaryID";
            temporaryID.Values = new String[] { TemporaryID };

            // LabelType
            DocumentManagement.StringValue labelType = new DocumentManagement.StringValue();
            labelType.Key = "LabelType";
            labelType.Values = new String[] { LabelType };

            // Client_ID
            DocumentManagement.IntegerValue clientID = new DocumentManagement.IntegerValue();
            clientID.Key = "Client_ID";
            clientID.Values = new Int64?[] { ClientID };

            // NumberOfCopies
            DocumentManagement.IntegerValue numberOfCopies = new DocumentManagement.IntegerValue();
            numberOfCopies.Key = "NumberOfCopies";
            numberOfCopies.Values = new Int64?[] { NumberOfCopies };

            // NumberOfLabels
            DocumentManagement.IntegerValue numberOfLabels = new DocumentManagement.IntegerValue();
            numberOfLabels.Key = "NumberOfLabels";
            numberOfLabels.Values = new Int64?[] { NumberOfLabels };

            // NumberOfItems
            DocumentManagement.IntegerValue numberOfItems = new DocumentManagement.IntegerValue();
            numberOfItems.Key = "NumberOfItems";
            numberOfItems.Values = new Int64?[] { NumberOfItems };

            // generateLabel
            DocumentManagement.IntegerValue generateLabel = new DocumentManagement.IntegerValue();
            generateLabel.Key = "generateLabel";
            if (GenerateLabel)
            {
                generateLabel.Values = new Int64?[] { 1 };
            }
            else
            {
                generateLabel.Values = new Int64?[] { 0 };
            }

            // poFromDate
            DocumentManagement.DateValue fromDate = new DocumentManagement.DateValue();
            fromDate.Key = "poFromDate";
            fromDate.Values = new DateTime?[] { FromDate };

            // poToDate
            DocumentManagement.DateValue toDate = new DocumentManagement.DateValue();
            toDate.Key = "poToDate";
            toDate.Values = new DateTime?[] { ToDate };

            // create the attribute group
            DocumentManagement.AttributeGroup poAtts = new DocumentManagement.AttributeGroup();
            poAtts.Key = "POCreateInfo";
            poAtts.Type = "POCreateInfo";
            poAtts.Values = new DocumentManagement.DataValue[] { subType, homeLocation, uniqueID, keywords, locatorType, refRate, offsiteStorageID, clientName, temporaryID, labelType, clientID,
                numberOfCopies, numberOfLabels, numberOfItems, generateLabel}; // , fromDate, toDate}; - no from and to dates at the moment - they need to be converted to a string, but what format? 

            return poAtts;

        }

        internal Boolean AssignToBox(Int64 ItemID, Int64 BoxID, Boolean UpdateLocation, Boolean UpdateRSI, Boolean UpdateStatus)
        {
            PhysicalObjects.PhysObjBoxingInfo boxInfo = new PhysicalObjects.PhysObjBoxingInfo();
            boxInfo.BoxID = BoxID;
            boxInfo.UpdateLocation = UpdateLocation;
            boxInfo.UpdateRSI = UpdateRSI;
            boxInfo.UpdateStatus = UpdateStatus;
            return poClient.PhysObjAssignToBox(ref poAuth, ItemID, boxInfo);
        }

        #endregion

        #endregion

        #region Helper methods

        public List<Int64> GetChildren(Int64 NodeID)
        {
            List<Int64> response = new List<Int64>();

            Node item = docClient.GetNode(ref docAuth, NodeID);
            if (item.IsContainer && item.ContainerInfo.ChildCount > 0)
            {
                GetNodesInContainerOptions getNodeOpts = new GetNodesInContainerOptions();
                getNodeOpts.MaxResults = 2000;
                getNodeOpts.MaxDepth = 0;
                Node[] nodes = docClient.GetNodesInContainer(ref docAuth, NodeID, getNodeOpts);
                for (int i = 0; i < nodes.Length; i++)
                {
                    response.Add(nodes[i].ID);
                }
            }

            return response;
        }

        internal void UpdateProjectFromTemplate(Int64 ProjectID, Int64 TemplateID)
        {
            CopyProjectParticipants(ProjectID, TemplateID);
            CopyChildren(ProjectID, TemplateID, ObjectType.Project);
        }

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
