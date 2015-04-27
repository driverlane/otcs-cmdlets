using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using cscmdlets.DocumentManagement;

namespace cscmdlets
{

    internal class SoapApi : CSMetadata
    {

        /* There's minimal error capture in this class.  This is deliberate.
         * Handle the error in the calling cmdlet method unless it's necessary at this level - and it sometimes is. */

        private CSClients server;
        internal List<Exception> NonTerminatingExceptions = new List<Exception>();

        internal SoapApi()
        {
            server = new CSClients(Globals.ServicesDirectory);
            server.AddAuthenticationDetails(Globals.Username, Globals.Password);
        }

        internal void CloseClients()
        {
            server.CloseClients();
        }

        #region API calls

        #region Core object calls

        internal Int64 CreateContainer(String Name, Int64 ParentID, Globals.ObjectType ObjectType)
        {
            Int64 newNode = 0;

            switch (ObjectType)
            {
                case Globals.ObjectType.Folder:
                    try
                    {
                        // open/check the client
                        if (server.docClient == null)
                            server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
                        else
                            server.CheckSession();

                        // create the folder
                        newNode = server.docClient.CreateSimpleFolder(ref server.docAuth, ParentID, Name);
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                DocumentManagement.Node node = server.docClient.GetNodeByName(ref server.docAuth, ParentID, Name);
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

                case Globals.ObjectType.Project:
                    try
                    {
                        // open/check the client
                        if (server.collabClient == null)
                            server.OpenClient(typeof(Collaboration.CollaborationClient));
                        else
                            server.CheckSession();

                        Collaboration.ProjectInfo projectInfo = new Collaboration.ProjectInfo();
                        projectInfo.Name = Name;
                        projectInfo.ParentID = ParentID;
                        Collaboration.Node node = server.collabClient.CreateProject(ref server.collabAuth, projectInfo);
                        newNode = node.ID;
                    }
                    catch (Exception e)
                    {
                        if (e.Message.EndsWith("already exists."))
                        {
                            try
                            {
                                // open/check the client
                                if (server.docClient == null)
                                    server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
                                else
                                    server.CheckSession();

                                DocumentManagement.Node node = server.docClient.GetNodeByName(ref server.docAuth, ParentID, Name);
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

        internal Int64 CreateSimpleDoc(String Name, Int64 ParentID, String Document)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // open/check the content client
            if (server.contClient == null)
                server.OpenClient(typeof(ContentService.ContentServiceClient));

            // get the context
            String context = server.docClient.CreateSimpleDocumentContext(ref server.docAuth, ParentID, Name);

            // set up the document details
            FileInfo file = new FileInfo(Document);
            ContentService.FileAtts fileAtts = new ContentService.FileAtts();
            fileAtts.CreatedDate = file.CreationTime;
            fileAtts.ModifiedDate = file.LastWriteTime;
            fileAtts.FileName = file.Name;
            fileAtts.FileSize = file.Length;
            FileStream fileStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);

            // upload the document
            String response = server.contClient.UploadContent(ref server.contAuth, context, fileAtts, fileStream).ToString();
            return Convert.ToInt64(response);
        }

        internal Int64 CreateRendition(Int64 NodeID, Int64 Version, String RenditionType, String Document)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // open/check the content client
            if (server.contClient == null)
                server.OpenClient(typeof(ContentService.ContentServiceClient));

            // get the context
            String context = server.docClient.CreateRenditionContext(ref server.docAuth, NodeID, Version, RenditionType);

            // set up the document details
            FileInfo file = new FileInfo(Document);
            ContentService.FileAtts fileAtts = new ContentService.FileAtts();
            fileAtts.CreatedDate = file.CreationTime;
            fileAtts.ModifiedDate = file.LastWriteTime;
            fileAtts.FileName = file.Name;
            fileAtts.FileSize = file.Length;
            FileStream fileStream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);

            // upload the document
            String response = server.contClient.UploadContent(ref server.contAuth, context, fileAtts, fileStream).ToString();
            return Convert.ToInt64(response);
        }

        internal void DeleteNode(Int64 NodeID)
        {
            // open/check the client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // delete the item
            server.docClient.DeleteNode(ref server.docAuth, NodeID);
        }

        #endregion

        #region Cats and atts

        internal List<String> ListNodeCategories(Int64 NodeID, Boolean ShowKey)
        {

            // open/check the client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the details for the item
            Node item = server.docClient.GetNode(ref server.docAuth, NodeID);

            // iterate through any categories
            List<String> cats = new List<string>();
            if (item.Metadata.AttributeGroups != null && item.Metadata.AttributeGroups.Length > 0)
            {
                String response = "";
                for (int i = 0; i < item.Metadata.AttributeGroups.Length; i++)
                {
                    if (ShowKey)
                    {
                        response = String.Format(",{0} - {1}", item.Metadata.AttributeGroups[i].DisplayName, item.Metadata.AttributeGroups[i].Key);
                    }
                    else
                    {
                        response = String.Format(",{0}", item.Metadata.AttributeGroups[i].DisplayName);
                    }
                    response = response.Substring(1);
                    cats.Add(response);
                }
            }
            else
            {
                cats.Add("No categories");
            }

            return cats;

        }

        internal Dictionary<String, List<Object>> ListAttributes(Int64 NodeID, Int64 CategoryID)
        {
            // open/check the client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            Dictionary<String, List<Object>> response = new Dictionary<string, List<Object>>();

            // get the item and check for a category
            Node item = server.docClient.GetNode(ref server.docAuth, NodeID);
            AttributeGroup cat = GetCategory(item.Metadata, CategoryID);

            if (cat != null)
            {
                for (int i = 0; i < cat.Values.Length; i++)
                {
                    // check if it's a table att and iterate the rows
                    if (cat.Values[i].GetType().Equals(typeof(TableValue)))
                    {
                        TableValue setAtt = (TableValue)cat.Values[i];
                        if (setAtt.Values != null)
                        {
                            for (int j = 0; j < setAtt.Values.Length; j++)
                            {
                                // iterate the atts on the row
                                List<DataValue> rowAtts = new List<DataValue>();
                                for (int k = 0; k < setAtt.Values[j].Values.Length; k++)
                                {
                                    // check the set row level attribute
                                    response.Add(String.Format("{0} - {1}.{2}", setAtt.Values[j].Values[k].Description, setAtt.Values[j].Values[k].Key, j + 1), (List<Object>)GetAttributeValues(setAtt.Values[j].Values[k]));
                                }
                            }
                        }
                    }
                    else
                        // standard attribute, so add the description and values
                        response.Add(String.Format("{0} - {1}", cat.Values[i].Description, cat.Values[i].Key), (List<Object>)GetAttributeValues(cat.Values[i]));
                }
            }
            else
                response.Add("No categories on item", null);

            return response;
        }

        internal void AddCategoryToNode(Int64 NodeID, Int64 CategoryID, Boolean Replace, Boolean MergeAttributes, Boolean UseNewValues)
        {
            // open/check the client

            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the node and category template
            Node item = server.docClient.GetNode(ref server.docAuth, NodeID);
            AttributeGroup cat = server.docClient.GetCategoryTemplate(ref server.docAuth, CategoryID);

            // add the category and clean any null strings
            AddCategory(item.Metadata, cat, Replace, MergeAttributes, UseNewValues);
            item.Metadata = CleanMetadata(item.Metadata);

            // update the item with the new metadata object
            server.docClient.UpdateNode(ref server.docAuth, item);
        }

        #endregion

        #region Member calls

        internal Int64 CreateUser(String Login, Int64 DepartmentGroupID, String Password, String FirstName, String MiddleName, String LastName, String Email, String Fax, String OfficeLocation,
            String Phone, String Title, Boolean LoginEnabled, Boolean PublicAccessEnabled, Boolean CreateUpdateUsers, Boolean CreateUpdateGroups, Boolean CanAdministerUsers, Boolean CanAdministerSystem)
        {
            // open/check the client
            if (server.memberClient == null)
                server.OpenClient(typeof(MemberService.MemberServiceClient));
            else
                server.CheckSession();

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

            Int64 response = server.memberClient.CreateUser(ref server.memberAuth, user);
            return response;
        }

        internal String DeleteMember(Int64 MemberID)
        {
            // open/check the client
            if (server.memberClient == null)
                server.OpenClient(typeof(MemberService.MemberServiceClient));
            else
                server.CheckSession();

            // delete the user
            server.memberClient.DeleteMember(ref server.memberAuth, MemberID);
            return "Deleted";
        }

        internal Int64 GetUserIDByLoginName(String Login)
        {
            // open/check the client
            if (server.memberClient == null)
                server.OpenClient(typeof(MemberService.MemberServiceClient));
            else
                server.CheckSession();

            // get the ID
            MemberService.Member user = server.memberClient.GetMemberByLoginName(ref server.memberAuth, Login);
            return user.ID;
        }

        internal Int64 CreateGroup(String Name, Int64 LeaderID)
        {
            // open/check the client
            if (server.memberClient == null)
                server.OpenClient(typeof(MemberService.MemberServiceClient));
            else
                server.CheckSession();

            // create the group
            Int64 response = server.memberClient.CreateGroup(ref server.memberAuth, Name, LeaderID);
            return response;
        }

        internal String AddMemberToGroup(Int64 GroupID, Int64 MemberID)
        {
            // open/check the client
            if (server.memberClient == null)
                server.OpenClient(typeof(MemberService.MemberServiceClient));
            else
                server.CheckSession();

            // add the member
            server.memberClient.AddMemberToGroup(ref server.memberAuth, GroupID, MemberID);
            return "Added";
        }

        #endregion

        #region Classifications calls

        internal Boolean AddClassifications(Int64 NodeID, Int64[] ClassIDs)
        {
            // open/check the client
            if (server.classClient == null)
                server.OpenClient(typeof(Classifications.ClassificationsClient));
            else
                server.CheckSession();

            // apply the classificatoin
            return server.classClient.ApplyClassifications(ref server.classAuth, NodeID, ClassIDs);
        }

        #endregion

        #region RM calls

        internal Boolean AddRMClassification(Int64 NodeID, Int64 RMClassID)
        {
            // open/check the client
            if (server.rmClient == null)
                server.OpenClient(typeof(RecordsManagement.RecordsManagementClient));
            else
                server.CheckSession();

            // add the rm classification
            RecordsManagement.RMAdditionalInfo info = new RecordsManagement.RMAdditionalInfo();
            Int64[] otherIDs = { };
            return server.rmClient.RMApplyClassification(ref server.rmAuth, NodeID, RMClassID, info, otherIDs);
        }

        internal Boolean FinaliseRecord(Int64 NodeID)
        {
            // open/check the client
            if (server.rmClient == null)
                server.OpenClient(typeof(RecordsManagement.RecordsManagementClient));
            else
                server.CheckSession();

            // add the finalise the item
            return server.rmClient.RMDeclareRecord(ref server.rmAuth, NodeID);
        }

        #endregion

        #region Physical objects calls

        internal Int64 CreatePhysicalItem(String Name, Int64 ParentID, Globals.PhysicalItemTypes Type, Int64 SubType, String HomeLocation, String Description, String UniqueID,
            String Keywords, String LocatorType, String RefRate, String OffsiteStorageID, String ClientName, String TemporaryID, String LabelType, Int64 ClientID, Int64 NumberOfCopies,
            Int64 NumberOfLabels, Int64 NumberOfItems, Boolean GenerateLabel, DateTime? FromDate, DateTime? ToDate)
        {

            // open/check the client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

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

            return server.docClient.CreateNode(ref server.docAuth, newNode).ID;
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
            // open/check the client
            if (server.poClient == null)
                server.OpenClient(typeof(PhysicalObjects.PhysicalObjectsClient));
            else
                server.CheckSession();

            PhysicalObjects.PhysObjBoxingInfo boxInfo = new PhysicalObjects.PhysObjBoxingInfo();
            boxInfo.BoxID = BoxID;
            boxInfo.UpdateLocation = UpdateLocation;
            boxInfo.UpdateRSI = UpdateRSI;
            boxInfo.UpdateStatus = UpdateStatus;
            return server.poClient.PhysObjAssignToBox(ref server.poAuth, ItemID, boxInfo);
        }

        #endregion

        #region Permissions

        internal String[] GetPermissions(Int64 NodeID, String[] Role, Int64 UserID)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // return the requested roles
            List<String> permlist = new List<string>();
            if (Role.Any("Owner".Equals) || Role.Any("All".Equals))
            {
                permlist.Add(String.Format("{0} - Owner: {1} Permissions: {2}", NodeID, perms.OwnerRight.RightID, GetPermissions(perms.OwnerRight)));
            }
            else if (Role.Any("OwnerGroup".Equals) || Role.Any("All".Equals))
            {
                permlist.Add(String.Format("{0} - OwnerGroup: {1} Permissions: {2}", NodeID, perms.OwnerGroupRight.RightID, GetPermissions(perms.OwnerGroupRight)));
            }
            else if (Role.Any("PublicAccess".Equals) || Role.Any("All".Equals))
            {
                permlist.Add(String.Format("{0} - Public Access Permissions: {1}", NodeID, GetPermissions(perms.PublicRight)));
            }
            else if (Role.Any("ACL".Equals) || Role.Any("All".Equals))
            {
                if (UserID > 0)
                {
                    try
                    {
                        DocumentManagement.NodeRight perm = perms.ACLRights.Single(item => item.RightID == UserID);
                        permlist.Add(String.Format("{0} - ACL: {1} Permissions: {2}", NodeID, perm.RightID, GetPermissions(perm)));
                    }
                    catch
                    {
                        permlist.Add(String.Format("{0} - ACL: {1} Permissions: Not assigned", NodeID, UserID));
                    }
                }
                else
                {
                    foreach (NodeRight perm in perms.ACLRights)
                    {
                        permlist.Add(String.Format("{0} - ACL: {1} Permissions: {2}", NodeID, perm.RightID, GetPermissions(perm)));
                    }
                }
            }

            return permlist.ToArray();

        }

        internal void SetOwner(Int64 NodeID, Int64 UserID, String[] Permissions)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // update the permissions, if specified
            if (Permissions != null && Permissions.Count() > 0)
                perms.OwnerRight = SetPermissions(perms.OwnerRight, Permissions);

            // update the owner, if specified
            if (UserID > 999)
                perms.OwnerRight.RightID = UserID;

            // update the permissions
            server.docClient.SetNodeRights(ref server.docAuth, NodeID, perms);

        }

        internal void SetOwnerGroup(Int64 NodeID, Int64 GroupID, String[] Permissions)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // update the permissions, if specified
            if (Permissions != null && Permissions.Count() > 0)
                perms.OwnerGroupRight = SetPermissions(perms.OwnerGroupRight, Permissions);

            // update the owner, if specified
            if (GroupID > 999)
                perms.OwnerGroupRight.RightID = GroupID;

            // update the permissions
            server.docClient.SetNodeRights(ref server.docAuth, NodeID, perms);

        }

        internal void SetPublicAccess(Int64 NodeID, String[] Permissions)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // update the permissions, if specified
            perms.PublicRight = SetPermissions(perms.PublicRight, Permissions);

            // update the permissions
            server.docClient.SetNodeRights(ref server.docAuth, NodeID, perms);

        }

        internal void SetAssignedAccess(Int64 NodeID, Int64 UserID, String[] Permissions)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // update the rights if the user/group is already assigned, otherwise add them
            try
            {
                DocumentManagement.NodeRight perm = perms.ACLRights.Single(item => item.RightID == UserID);
                perm = SetPermissions(perm, Permissions);
            }
            catch
            {
                // create a new ACL object
                DocumentManagement.NodeRight perm = new NodeRight();
                perm.RightID = UserID;
                perm.Permissions = new NodePermissions();
                perm.Type = "ACL";
                perm = SetPermissions(perm, Permissions);

                // add the new assigned access to the permissions object
                List<NodeRight> newPerms = new List<NodeRight>();
                if (perms.ACLRights != null)
                    newPerms = perms.ACLRights.ToList();
                newPerms.Add(perm);
                perms.ACLRights = newPerms.ToArray();
            }

            // update the permissions
            server.docClient.SetNodeRights(ref server.docAuth, NodeID, perms);

        }

        internal void RemoveAssignedAccess(Int64 NodeID, Int64 UserID)
        {
            // open/check the doc management client
            if (server.docClient == null)
                server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
            else
                server.CheckSession();

            // get the existing permissions on the object
            DocumentManagement.NodeRights perms = server.docClient.GetNodeRights(ref server.docAuth, NodeID);

            // check if the user/group is assigned and removed
            Boolean update = false;
            if (perms.ACLRights != null)
            {
                List<NodeRight> newPerms = new List<NodeRight>();
                foreach (NodeRight perm in perms.ACLRights)
                {
                    if (perm.RightID != UserID)
                    {
                        newPerms.Add(perm);
                    }
                    else
                        update = true;
                }

                if (update)
                {
                    if (newPerms.Count == 0)
                        perms.ACLRights = null;
                    else
                        perms.ACLRights = newPerms.ToArray();
                    server.docClient.SetNodeRights(ref server.docAuth, NodeID, perms);
                }
            }

        }

        #endregion

        #endregion

        #region Helper methods

        public List<Int64> GetChildren(Int64 NodeID)
        {
            List<Int64> response = new List<Int64>();

            Node item = server.docClient.GetNode(ref server.docAuth, NodeID);
            if (item.IsContainer && item.ContainerInfo.ChildCount > 0)
            {
                GetNodesInContainerOptions getNodeOpts = new GetNodesInContainerOptions();
                getNodeOpts.MaxResults = 2000;
                getNodeOpts.MaxDepth = 0;
                Node[] nodes = server.docClient.GetNodesInContainer(ref server.docAuth, NodeID, getNodeOpts);
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
            CopyChildren(ProjectID, TemplateID, Globals.ObjectType.Project);
        }

        private void CopyProjectParticipants(Int64 ProjectID, Int64 TemplateID)
        {

            try
            {
                // open/check the client
                if (server.collabClient == null)
                    server.OpenClient(typeof(Collaboration.CollaborationClient));
                else
                    server.CheckSession();

                // copy the public access
                Collaboration.ProjectInfo templateInfo = server.collabClient.GetProject(ref server.collabAuth, TemplateID);
                Collaboration.ProjectInfo projectInfo = server.collabClient.GetProject(ref server.collabAuth, ProjectID);
                projectInfo.PublicAccess = templateInfo.PublicAccess;

                // add the participants from the template
                Collaboration.ProjectParticipants templateParticipants = server.collabClient.GetParticipants(ref server.collabAuth, TemplateID);
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
                Collaboration.ProjectParticipants projectParticipants = server.collabClient.GetParticipants(ref server.collabAuth, ProjectID);
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
                server.collabClient.UpdateProjectParticipants(ref server.collabAuth, ProjectID, participants.ToArray());
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void CopyChildren(Int64 Copy, Int64 Template, Globals.ObjectType ObjectType)
        {
            try
            {
                // open/check the client
                if (server.docClient == null)
                    server.OpenClient(typeof(DocumentManagement.DocumentManagementClient));
                else
                    server.CheckSession();

                DocumentManagement.Node parentNode = server.docClient.GetNode(ref server.docAuth, Copy);
                DocumentManagement.Node newParentNode = server.docClient.GetNode(ref server.docAuth, Template);

                if (parentNode.IsContainer && newParentNode.IsContainer)
                {
                    // change the ID if it's a volume object
                    Boolean VolumeObject = false;
                    if (ObjectType.Equals(Globals.ObjectType.Project)) { VolumeObject = true; }
                    if (VolumeObject)
                    {
                        Copy = -Copy;
                        Template = -Template;
                    }

                    DocumentManagement.Node[] children = server.docClient.ListNodes(ref server.docAuth, Template, true);
                    if (children != null)
                    {
                        foreach (DocumentManagement.Node child in children)
                        {
                            DocumentManagement.CopyOptions copyOptions = new DocumentManagement.CopyOptions();
                            server.docClient.CopyNode(ref server.docAuth, child.ID, Copy, child.Name, copyOptions);
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

            String[] rootNodes = server.docClient.GetRootNodeTypes(ref server.docAuth);
            DocumentManagement.Node child = server.docClient.GetNode(ref server.docAuth, NodeID);
            DocumentManagement.Node parent = server.docClient.GetNode(ref server.docAuth, child.ParentID);
            String path = child.Name;
            path = parent.Name + " >> " + path;

            do
            {
                child = server.docClient.GetNode(ref server.docAuth, parent.ID);
                parent = server.docClient.GetNode(ref server.docAuth, child.ParentID);

                path = parent.Name + " >> " + path;

            } while (!rootNodes.Contains(parent.Name + "WS"));

            return path;
        }

        private String GetParentName(Int64 NodeID)
        {
            String name;
            DocumentManagement.Node child = server.docClient.GetNode(ref server.docAuth, NodeID);
            DocumentManagement.Node parent = server.docClient.GetNode(ref server.docAuth, child.ParentID);
            name = parent.Name;
            return name;
        }

        private String GetPermissions(NodeRight Permission)
        {
            String permlist = "No permissions";

            if (Permission.Permissions.SeePermission)
                permlist = "See";

            if (Permission.Permissions.SeeContentsPermission)
                permlist += ",SeeContents";

            if (Permission.Permissions.ModifyPermission)
                permlist += ",Modify";

            if (Permission.Permissions.EditAttributesPermission)
                permlist += ",EditAttributes";

            if (Permission.Permissions.AddItemsPermission)
                permlist += ",AddItems";

            if (Permission.Permissions.ReservePermission)
                permlist += ",Reserve";

            if (Permission.Permissions.DeleteVersionsPermission)
                permlist += ",DeleteVersions";

            if (Permission.Permissions.DeletePermission)
                permlist += ",Delete";

            if (Permission.Permissions.EditPermissionsPermission)
                permlist += ",EditPermissions";

            return permlist;
        }

        private DocumentManagement.NodeRight SetPermissions(NodeRight Permission, String[] Permissions)
        {

            Permissions = ExpandPermissions(Permissions);

            if (Permission.Permissions == null)
                Permission.Permissions = new NodePermissions();

            // set the permissions if it's in the Permissions array
            if (Permissions.Any("AddItems".Equals))
                Permission.Permissions.AddItemsPermission = true;
            else
                Permission.Permissions.AddItemsPermission = false;

            if (Permissions.Any("Delete".Equals))
                Permission.Permissions.DeletePermission = true;
            else
                Permission.Permissions.DeletePermission = false;

            if (Permissions.Any("DeleteVersions".Equals))
                Permission.Permissions.DeleteVersionsPermission = true;
            else
                Permission.Permissions.DeleteVersionsPermission = false;

            if (Permissions.Any("EditAttributes".Equals))
                Permission.Permissions.EditAttributesPermission = true;
            else
                Permission.Permissions.EditAttributesPermission = false;

            if (Permissions.Any("EditPermissions".Equals))
                Permission.Permissions.EditPermissionsPermission = true;
            else
                Permission.Permissions.EditPermissionsPermission = false;

            if (Permissions.Any("Modify".Equals))
                Permission.Permissions.ModifyPermission = true;
            else
                Permission.Permissions.ModifyPermission = false;

            if (Permissions.Any("Reserve".Equals))
                Permission.Permissions.ReservePermission = true;
            else
                Permission.Permissions.ReservePermission = false;

            if (Permissions.Any("SeeContents".Equals))
                Permission.Permissions.SeeContentsPermission = true;
            else
                Permission.Permissions.SeeContentsPermission = false;

            if (Permissions.Any("See".Equals))
                Permission.Permissions.SeePermission = true;
            else
                Permission.Permissions.SeePermission = false;

            return Permission;
        }

        private String[] ExpandPermissions(String[] Permissions)
        {
            List<String> newPerms = Permissions.ToList();

            // if they've got EditPermissions, they get everything
            if (Permissions.Any("EditPermissions".Equals))
                return new String[] { "See", "SeeContents", "Modify", "EditAttributes", "AddItems", "Reserve", "DeleteVersions", "Delete", "EditPermissions" };

            // if it's Delete it's got to also have See, SeeContents, Modify and DeleteVersions
            if (newPerms.Any("Delete".Equals))
            {
                if (!newPerms.Any("DeleteVersions".Equals))
                    newPerms.Add("DeleteVersions");
                if (!newPerms.Any("Modify".Equals))
                    newPerms.Add("Modify");
                if (!newPerms.Any("SeeContents".Equals))
                    newPerms.Add("SeeContents");
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's DeleteVersions it's got to also have See, SeeContents and Modify
            if (newPerms.Any("DeleteVersions".Equals))
            {
                if (!newPerms.Any("Modify".Equals))
                    newPerms.Add("Modify");
                if (!newPerms.Any("SeeContents".Equals))
                    newPerms.Add("SeeContents");
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's Reserve it's got to also have See, SeeContents and Modify
            if (newPerms.Any("Reserve".Equals))
            {
                if (!newPerms.Any("Modify".Equals))
                    newPerms.Add("Modify");
                if (!newPerms.Any("SeeContents".Equals))
                    newPerms.Add("SeeContents");
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's AddItems it's got to also have See and Modify
            if (newPerms.Any("AddItems".Equals))
            {
                if (!newPerms.Any("Modify".Equals))
                    newPerms.Add("Modify");
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's EditAttributes it's got to also have See, SeeContents and Modify
            if (newPerms.Any("EditAttributes".Equals))
            {
                if (!newPerms.Any("Modify".Equals))
                    newPerms.Add("Modify");
                if (!newPerms.Any("SeeContents".Equals))
                    newPerms.Add("SeeContents");
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's Modify it's got to also have See
            if (newPerms.Any("EditAttributes".Equals))
            {
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            // if it's SeeContents it's got to also have See
            if (newPerms.Any("SeeContents".Equals))
            {
                if (!newPerms.Any("See".Equals))
                    newPerms.Add("See");
                return newPerms.ToArray();
            }

            return newPerms.ToArray();
        }

        #endregion

    }

}
