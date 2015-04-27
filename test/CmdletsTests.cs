using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using cscmdlets;

namespace cscmdlets.tests
{

    public class TestGlobals
    {

        public static readonly Dictionary<string, object> current = new Dictionary<string, object>{

            // environment details
            {"UserName", "admin"},
            {"Password",  "p@ssw0rd"},
            {"ServicesDirectory", "http://cgi-eim.cloudapp.net/cws/"},
            {"RestUrl", "http://cgi-eim.cloudapp.net/otcs/cs.exe/api/v1/"},

            // test folder
            {"ParentID", 116609},

            // project workspace copying
            {"MasterWorkspaceID", 116389},
            {"TemplateWorkspaceID", 116499},

            // cats and atts testing
            {"Cat1ID", 116500},
            {"Cat1Name", "cscmdlets1"},
            {"Cat1Version", 2},
            {"Cat2ID", 116501},
            {"Cat2Name", "cscmdlets2"},
            {"Cat2Version", 1},

            // user and group testing
            {"DepartmentGroupID", 1001},
            {"UserToRetrieve", "admin"},
            {"UserToRetrieveID", 1000},

            // classifications testing
            {"ClassificationIDs", "116281,115627"},
            {"RMClassificationID", 116174},

            // physical objects testing
            {"ItemSubType", 116060},
            {"PartSubType", 116503},
            {"BoxSubType", 116504},
            {"HomeLocation", "Compactus Level 1 Building 3"},

            // document upload
            {"DocPath", "D:\\code\\cscmdlets\\test\\tester.docx"},

            // permissions
            {"Permissions", "See,SeeContents"},
            {"OtherUser", 1002},
            {"OtherGroup", 250347},
            {"NewPermissions", "See,SeeContents,Modify"},

            // template workspaces
            {"TemplateWorkspacesParentId", 256177},
            {"BinderTemplateID", 256947},
            {"BinderClassificationID", 264648},
            {"CaseTemplateID", 264098},
            {"CaseClassificationID", 264097}

        };

        public static readonly Dictionary<string, object> cs10 = new Dictionary<string, object>{

            // environment details
            {"UserName", "admin"},
            {"Password","p@ssw0rd"},
            {"ServicesDirectory","http://content2.cgi.demo/les-services/"},

            // test folder
            {"ParentID", 23996},

            // project workspace copying
            {"MasterWorkspaceID", 26195},
            {"TemplateWorkspaceID", 25975},

            // cats and atts testing
            {"Cat1ID", 26086},
            {"Cat1Name", "Document"},
            {"Cat1Version", 2},
            {"Cat2ID", 26087},
            {"Cat2Name", "Drawing"},
            {"Cat2Version", 1},

            // user and group testing
            {"DepartmentGroupID", 1001},
            {"UserToRetrieve", "admin"},
            {"UserToRetrieveID", 1000},

            // classifications testing
            {"ClassificationIDs", "25866,25976"},
            {"RMClassificationID", 25647},

            // physical objects testing
            {"ItemSubType", 25977},
            {"PartSubType", 25779},
            {"BoxSubType", 26219},
            {"HomeLocation", "Compactus Level 1 Building 3"},

            // document upload
            {"DocPath", "C:\\code\\cscmdlets\\test\\tester.docx"},

            // permissions
            {"Permissions", "See,SeeContents"},
            {"NewPermissions", "See,SeeContents,Modify"}
        };

    }

    public abstract class PSTestFixture
    {

        protected Runspace runspace { get; set; }

        public void CreateRunspace()
        {
            runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            String cmd = @"Import-Module .\cscmdlets.dll";
            ExecPS(cmd);

        }

        public void CloseRunspace()
        {
            runspace.Close();
            runspace.Dispose();
        }

        protected Collection<PSObject> ExecPS(String Command)
        {
            return runspace.CreatePipeline(Command).Invoke();
        }

        protected void OpenConnections()
        {
            String cmd = String.Format("Open-CSConnections -Username {0} -Password {1} -ServicesDirectory {2} -Url {3}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["ServicesDirectory"], TestGlobals.current["RestUrl"]);
            ExecPS(cmd);
        }

        public String UniqueName()
        {
            return String.Format("Testing{0}", DateTime.Now.ToString("yyyymmddHHmm"));
        }

    }

    #region Encryption

    [TestClass]
    public class ConvertToEncryptedPasswordCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        // todo later test the encryption and logging in with encrypted password
    }

    #endregion

    #region Connection

    [TestClass]
    public class OpenCSConnectionsCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        public void OpenConnectionsTest()
        {
            // arrange
            String cmd = String.Format("Open-CSConnections -Username {0} -Password {1} -ServicesDirectory {2} -Url {3}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["ServicesDirectory"], TestGlobals.current["RestUrl"]);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)

                .First();

            // assert
            Assert.AreEqual("Connections established", result);
        }

    }

    [TestClass]
    public class OpenCSConnectionCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        public void OpenSoapConnection()
        {
            // arrange
            String cmd = String.Format("Open-CSConnectionSOAP -Username {0} -Password {1} -ServicesDirectory {2}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["ServicesDirectory"]);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)

                .First();

            // assert
            Assert.AreEqual("Connection established", result);
        }

    }

    [TestClass]
    public class OpenCSConnectionRestCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        public void OpenRestConnection()
        {
            // arrange
            String cmd = String.Format("Open-CSConnectionREST -Username {0} -Password {1} -Url {2}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["RestUrl"]);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)

                .First();

            // assert
            Assert.AreEqual("Connection established", result);
        }

    }

    #endregion

    #region Core objects

    [TestClass]
    public class AddCSProjectWorkspaceCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddProjectWorkspace()
        {
            // arrange
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)

                .First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddProjectWorkspace()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a project workspace
            Assert.AreEqual(result, result2);                   // the second run returned the ID from the first run
        }

        [TestMethod]
        public void AddProjectWorkspaceWithTemplate()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1} -TemplateID {2}", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["MasterWorkspaceID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a project workspace with template
            Assert.AreEqual(result, result2);                   // the second run returned the ID from the first run
        }

    }

    [TestClass]
    public class AddCSFolderCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddFolder()
        {
            // arrange
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddFolder()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a folder
            Assert.AreEqual(result, result2);                   // the second run returned the ID from the first run

        }

    }

    [TestClass]
    public class AddCSDocumentCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddDocument()
        {
            // arrange
            String cmd = String.Format("Add-CSDocument -Name {0} -ParentID {1} -Document {2}", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["DocPath"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddDocument()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSDocument -Name {0} -ParentID {1} -Document {2}", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["DocPath"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a document

        }

    }

    [TestClass]
    public class RemoveCSNodeCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionRemoveNode()
        {
            // arrange
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveNode()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - deleted", result), result2);
        }

    }

    [TestClass]
    public class CatsAndAttsTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionGetCategories()
        {
            // arrange
            String cmd = String.Format("Get-CSCategories -NodeID {0}", TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddCategory()
        {
            // arrange
            String cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {0}", TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void GetCategories()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var item = ExecPS(cmd).Select(o => o.BaseObject).First();

            String cmd1 = String.Format("Get-CSCategories -NodeID {0}", item);
            String cmd2 = String.Format("Get-CSCategories -NodeID {0} -ShowKey", item);

            // act
            List<String> result1 = (List<String>)ExecPS(cmd1).Select(o => o.BaseObject).First();
            List<String> result2 = (List<String>)ExecPS(cmd2).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - No categories", item), result1.First());
            Assert.AreEqual(String.Format("{0} - No categories", item), result2.First());
        }

        [TestMethod]
        public void GetAttributeValues()
        {
            OpenConnections();

            // get the attributes on the empty folder
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var item = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSAttributeValues -NodeID {0} -CategoryID {1}", item, TestGlobals.current["Cat1ID"]);
            Dictionary<String, List<Object>> result1 = (Dictionary<String, List<Object>>)ExecPS(cmd).Select(o => o.BaseObject).First();

            // get the attributes on the populated folder
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.current["Cat1ID"]);
            ExecPS(cmd);
            cmd = String.Format("Get-CSAttributeValues -NodeID {0} -CategoryID {1}", item, TestGlobals.current["Cat1ID"]);
            Dictionary<String, List<Object>> result2 = (Dictionary<String, List<Object>>)ExecPS(cmd).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(1, result1.Count);
            Assert.AreEqual("No categories on item", result1.ElementAt(0).Key);
            Assert.AreEqual(3, result2.Count);
            Assert.AreEqual(String.Format("Field 3 - {0}.2.4", TestGlobals.current["Cat1ID"]), result2.ElementAt(2).Key);
        }

        [TestMethod]
        public void AddCategory()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var item = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.current["Cat1ID"]);
            ExecPS(cmd);
            String cmd1 = String.Format("Get-CSCategories -NodeID {0}", item);
            String cmd2 = String.Format("Get-CSCategories -NodeID {0} -ShowKey", item);

            // act
            List<String> result1 = (List<String>)ExecPS(cmd1).Select(o => o.BaseObject).First();
            List<String> result2 = (List<String>)ExecPS(cmd2).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.current["Cat2ID"]);
            ExecPS(cmd);
            List<String> result3 = (List<String>)ExecPS(cmd1).Select(o => o.BaseObject).First();
            List<String> result4 = (List<String>)ExecPS(cmd2).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.current["Cat1Name"]), result1.First());
            Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.current["Cat1Name"], TestGlobals.current["Cat1ID"], TestGlobals.current["Cat1Version"]), result2.First());
            Assert.AreEqual(2, result3.Count);
            foreach (String cat in result3)
            {
                if (cat.Contains(TestGlobals.current["Cat1Name"].ToString()))
                {
                    Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.current["Cat1Name"]), cat);
                }
                else
                {
                    Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.current["Cat2Name"]), cat);
                }

            }
            Assert.AreEqual(2, result4.Count);
            foreach (String cat in result4)
            {
                if (cat.Contains(TestGlobals.current["Cat1Name"].ToString()))
                {
                    Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.current["Cat1Name"], TestGlobals.current["Cat1ID"], TestGlobals.current["Cat1Version"]), cat);
                }
                else
                {
                    Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.current["Cat2Name"], TestGlobals.current["Cat2ID"], TestGlobals.current["Cat2Version"]), cat);
                }

            }

        }

    }

    #endregion

    #region Members

    [TestClass]
    public class AddCSUserCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Convert.ToInt64(User) > 0)
            {
                String cmd = String.Format("Remove-CSMember -MemberID {0}", Convert.ToInt64(User));
                ExecPS(cmd);
            }
            CloseRunspace();

        }

        object User;

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddUser()
        {
            // arrange
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void CreateUser()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            String.Format("Remove-CSMember -MemberID {0}", result);

            // assert

            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        public void CreateExistingUser()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            User = result;

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - user NOT created. ERROR: Error creating a new user. [E662437890]", UniqueName()), result2);
        }

    }

    [TestClass]
    public class AddCSGroupCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            if (Convert.ToInt64(Group) > 0)
            {
                String cmd = String.Format("Remove-CSMember -MemberID {0}", Group);
                ExecPS(cmd);
            }
            CloseRunspace();

        }

        object Group;

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddGroup()
        {
            // arrange
            String cmd = String.Format("Add-CSGroup -Name {0} -LeaderID {1}", UniqueName(), TestGlobals.current["OtherGroup"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void CreateGroup()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSGroup -Name {0}", UniqueName());

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            String.Format("Remove-CSMember -MemberID {0}", result);

            // assert

            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        public void CreateGroupWithLeader()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSGroup -Name {0} -LeaderID {1}", UniqueName(), TestGlobals.current["OtherGroup"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            String.Format("Remove-CSMember -MemberID {0}", result);

            // assert

            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        public void CreateExistingGroup()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSGroup -Name {0}", UniqueName());
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            Group = result;

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - group NOT created. ERROR: Could not create group '{0}': 'Specified name already exists.'. [E662437890]", UniqueName()), result2);
        }

    }

    [TestClass]
    public class RemoveCSMemberCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionRemoveUser()
        {
            // arrange
            String cmd = "Remove-CSMember -MemberID 1";

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveUser()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSMember -MemberID {0}", result);
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - Deleted", result), result2);
        }

    }

    [TestClass]
    public class GetCSUserIDByLoginCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {

        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionGetUserID()
        {
            // arrange
            String cmd = String.Format("Get-CSUserIDByLogin -Login {0}", TestGlobals.current["UserToRetrieve"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void GetUserID()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSUserIDByLogin -Login {0}", TestGlobals.current["UserToRetrieve"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(result, Convert.ToInt64(TestGlobals.current["UserToRetrieveID"]));
        }

    }

    #endregion

    #region Classifications

    [TestClass]
    public class AddCSClassificationsCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddClassifications()
        {
            // arrange
            String cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", TestGlobals.current["ParentID"], TestGlobals.current["ClassificationIDs"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddClassifications()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", result, TestGlobals.current["ClassificationIDs"]);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - classifications applied", result), result2);
        }

    }

    #endregion

    #region Records management

    [TestClass]
    public class AddCSRMClassificationCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddRMClassification()
        {
            // arrange
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", TestGlobals.current["ParentID"], TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddRMClassification()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.current["RMClassificationID"]);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - RM classification applied", result), result2);
        }

    }

    [TestClass]
    public class SetCSFinaliseRecordCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionFinaliseRecord()
        {
            // arrange
            String cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void FinaliseRecord()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.current["RMClassificationID"]);
            ExecPS(cmd);
            cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", result);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - item finalised", result), result2);
        }

    }

    #endregion

    #region Physical objects

    [TestClass]
    public class AddCSPhysItemCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddPhysItem()
        {
            // arrange
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysItem()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
            Assert.IsInstanceOfType(result, typeof(Int64));
        }

    }

    [TestClass]
    public class AddCSPhysContainerCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddPhysContainer()
        {
            // arrange
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["PartSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysContainer()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["PartSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));
        }

    }

    [TestClass]
    public class AddCSPhysBoxCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddPhysBox()
        {
            // arrange
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["BoxSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysBox()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSPhysBox -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["BoxSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));
        }

    }

    [TestClass]
    public class SetCSPhysObjToBoxCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAssignToBox()
        {
            // arrange
            String cmd = String.Format("Set-CSPhysObjToBox -ItemID {0} -BoxID {0}", TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AssignToBox()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["PartSubType"], TestGlobals.current["HomeLocation"]);
            var item = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSPhysBox -Name {0}box -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);
            var box = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Set-CSPhysObjToBox -ItemID {0} -BoxID {1}", item, box);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);
            cmd = String.Format("Remove-CSNode -NodeID {0}", box);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - assigned to box {1}", item, box), result);
        }

    }

    #endregion

    #region Permissions

    [TestClass]
    public class GetCSPermissions : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionGetPermissions()
        {
            // arrange
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role {1}", TestGlobals.current["ParentID"], "All");

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void GetPermissions()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role {1}", TestGlobals.current["ParentID"], "Owner");

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - Owner: {1} Permissions: {2}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"], TestGlobals.current["Permissions"]), result);
        }

    }

    [TestClass]
    public class SetCSOwner : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionSetOwner()
        {
            // arrange
            String cmd = String.Format("Set-CSOwner -NodeID {0} -UserID {1} -Permissions {2}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"], TestGlobals.current["Permissions"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void SetOwnerPerms()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role Owner", TestGlobals.current["ParentID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSOwner -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["NewPermissions"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role Owner", TestGlobals.current["ParentID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Set-CSOwner -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["Permissions"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - owner updated", TestGlobals.current["ParentID"]), output);
            Assert.AreNotEqual(initialResults.First(), updatedResults.First(), "New permissions assigned to owner");
        }

        [TestMethod]
        public void SetOwner()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role Owner", TestGlobals.current["ParentID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSOwner -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["OtherUser"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role Owner", TestGlobals.current["ParentID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Set-CSOwner -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - owner updated", TestGlobals.current["ParentID"]), output);
            Assert.AreNotEqual(initialResults.First(), updatedResults.First(), "New user assigned to owner");
        }

    }

    [TestClass]
    public class SetCSOwnerGroup : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionSetOwnerGroup()
        {
            // arrange
            String cmd = String.Format("Set-CSOwnerGroup -NodeID {0} -GroupID {1} -Permissions {2}", TestGlobals.current["ParentID"], TestGlobals.current["DepartmentGroupID"], TestGlobals.current["Permissions"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void SetOwnerGroupPerms()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role OwnerGroup", TestGlobals.current["ParentID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSOwnerGroup -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["NewPermissions"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role OwnerGroup", TestGlobals.current["ParentID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Set-CSOwnerGroup -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["Permissions"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - owner group updated", TestGlobals.current["ParentID"]), output);
            Assert.AreNotEqual(initialResults.First(), updatedResults.First(), "New permissions assigned to OwnerGroup");
        }

        [TestMethod]
        public void SetOwnerGroup()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role OwnerGroup", TestGlobals.current["ParentID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSOwnerGroup -NodeID {0} -GroupID {1}", TestGlobals.current["ParentID"], TestGlobals.current["OtherGroup"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role OwnerGroup", TestGlobals.current["ParentID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Set-CSOwnerGroup -NodeID {0} -GroupID {1}", TestGlobals.current["ParentID"], TestGlobals.current["DepartmentGroupID"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - owner group updated", TestGlobals.current["ParentID"]), output);
            Assert.AreNotEqual(initialResults.First(), updatedResults.First(), "New group assigned to OwnerGroup");
        }

    }

    [TestClass]
    public class SetCSPublicAccess : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionSetPublicAccess()
        {
            // arrange
            String cmd = String.Format("Set-CSPublicAccess -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["Permissions"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void SetPublicAccess()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Get-CSPermissions -NodeID {0} -Role PublicAccess", TestGlobals.current["ParentID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSPublicAccess -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["NewPermissions"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role PublicAccess", TestGlobals.current["ParentID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Set-CSPublicAccess -NodeID {0} -Permissions {1}", TestGlobals.current["ParentID"], TestGlobals.current["Permissions"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - public access updated", TestGlobals.current["ParentID"]), output);
            Assert.AreNotEqual(initialResults.First(), updatedResults.First(), "New permissions assigned to PublicAccess");
        }

    }

    [TestClass]
    public class SetCSAssignedAccess : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionSetAssignedAccess()
        {
            // arrange
            String cmd = String.Format("Set-CSAssignedAccess -NodeID {0} -UserID {1} -Permissions {2}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"], TestGlobals.current["Permissions"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void SetAssignedAccess()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Remove-CSAssignedAccess -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            ExecPS(cmd);
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role ACL -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Set-CSAssignedAccess -NodeID {0} -UserID {1} -Permissions {2}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"], TestGlobals.current["Permissions"]);

            // act
            var output = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role ACL -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // clean up
            cmd = String.Format("Remove-CSAssignedAccess -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - permissions applied", TestGlobals.current["ParentID"]), output);
            Assert.IsTrue(initialResults.First().ToString().Contains("Not assigned"), "User is not assigned to object");
            Assert.IsTrue(!updatedResults.First().ToString().Contains("Not assigned"), "User has been assigned to object");
        }

    }

    [TestClass]
    public class RemoveCSAssignedAccess : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionRemoveAssignedAccess()
        {
            // arrange
            String cmd = String.Format("Remove-CSAssignedAccess -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveAssignedAccess()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Set-CSAssignedAccess -NodeID {0} -UserID {1} -Permissions {2}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"], TestGlobals.current["Permissions"]);
            ExecPS(cmd);
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role ACL -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            var initialResults = ExecPS(cmd).Select(o => o.BaseObject);
            cmd = String.Format("Remove-CSAssignedAccess -NodeID {0} -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Get-CSPermissions -NodeID {0} -Role ACL -UserID {1}", TestGlobals.current["ParentID"], TestGlobals.current["UserToRetrieveID"]);
            var updatedResults = ExecPS(cmd).Select(o => o.BaseObject);

            // assert
            Assert.AreEqual(String.Format("{0} - permissions removed", TestGlobals.current["ParentID"]), result);
            Assert.IsTrue(!initialResults.First().ToString().Contains("Not assigned"), "User is assigned to object");
            Assert.IsTrue(updatedResults.First().ToString().Contains("Not assigned"), "User has been removed from object");
        }

    }

    #endregion

    #region Template Workspaces

    [TestClass]
    public class AddCSBinderCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddBinder()
        {
            // arrange
            String cmd = String.Format("Add-CSBinder -TemplateID {0} -ParentID {1} -ClassificationID {2} -Name {3} -Description {4}", TestGlobals.current["BinderTemplateID"], TestGlobals.current["TemplateWorkspacesParentId"], TestGlobals.current["BinderClassificationID"], UniqueName(), "Description");

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddBinder()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSBinder -TemplateID {0} -ParentID {1} -ClassificationID {2} -Name {3} -Description {4}", TestGlobals.current["BinderTemplateID"], TestGlobals.current["TemplateWorkspacesParentId"], TestGlobals.current["BinderClassificationID"], UniqueName(), "Description");

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a binder

        }

        [TestMethod]
        public void AddCase()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSBinder -TemplateID {0} -ParentID {1} -ClassificationID {2} -Name {3} -Description {4}", TestGlobals.current["BinderTemplateID"], TestGlobals.current["TemplateWorkspacesParentId"], TestGlobals.current["BinderClassificationID"], UniqueName(), "Description");
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSCase -TemplateID {0} -ParentID {1} -ClassificationID {2} -Name {3} -Description {4}", TestGlobals.current["CaseTemplateID"], result, TestGlobals.current["CaseClassificationID"], UniqueName(), "Description");

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result2, typeof(Int64));     // created a binder

        }

    }

    #endregion

    #region Renditions

    [TestClass]
    public class AddCSRenditionCommandTests : PSTestFixture
    {
        [TestInitialize]
        public void Setup()
        {
            CreateRunspace();
        }

        [TestCleanup]
        public void TearDown()
        {
            CloseRunspace();
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddRendition()
        {
            // arrange
            String cmd = String.Format("Add-CSRendition -NodeID {0} -Version {1} -Type {2} -Document {3}", TestGlobals.current["ParentID"], 1, "PDF", TestGlobals.current["DocPath"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddRendition()
        {
            // arrange
            OpenConnections();
            String cmd = String.Format("Add-CSDocument -Name {0} -ParentID {1} -Document {2}", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["DocPath"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSRendition -NodeID {0} -Version {1} -Type {2} -Document {3}", Convert.ToInt64(result), 1, "PDF", TestGlobals.current["DocPath"]);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.AreEqual("Rendition added", result2);

        }

    }

    #endregion

}