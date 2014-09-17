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
            {"Password","p@ssw0rd"},
            {"ServicesDirectory","http://content.cgi.demo/cws/"},

            // test folder
            {"ParentID", 7944},

            // project workspace copying
            {"MasterWorkspaceID", 30262},
            {"TemplateWorkspaceID", 28611},

            // cats and atts testing
            {"Cat1ID", 28941},
            {"Cat1Name", "Document"},
            {"Cat1Version", 2},
            {"Cat2ID", 30154},
            {"Cat2Name", "Drawing"},
            {"Cat2Version", 1},

            // user and group testing
            {"DepartmentGroupID", 1001},
            {"UserToRetrieve", "admin"},
            {"UserToRetrieveID", 1000},

            // classifications testing
            {"ClassificationIDs", "32239,30699"},
            {"RMClassificationID", 31249},

            // physical objects testing
            {"ItemSubType", 31360},
            {"PartSubType", 30481},
            {"BoxSubType", 32019},
            {"HomeLocation", "Compactus Level 1 Building 3"},

            // document upload
            {"DocPath", "C:\\code\\cscmdlets\\test\\tester.docx"}
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
            {"DocPath", "C:\\code\\cscmdlets\\test\\tester.docx"}
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

        protected void OpenConnection()
        {
            String cmd = String.Format("Open-CSConnection -Username {0} -Password {1} -ServicesDirectory {2}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["ServicesDirectory"]);
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
        public void OpenTheConnection()
        {
            // arrange
            String cmd = String.Format("Open-CSConnection -Username {0} -Password {1} -ServicesDirectory {2}", TestGlobals.current["UserName"], TestGlobals.current["Password"], TestGlobals.current["ServicesDirectory"]);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)

                .First();

            // assert
            Assert.AreEqual("Connection established", result);
        }

    }

    #endregion

    #region Document management webservice

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
            OpenConnection();
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
            OpenConnection();
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
            OpenConnection();
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
            OpenConnection();
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
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(String.Format("{0} - Deleted", result), result2);
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
            OpenConnection();
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
            OpenConnection();

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
            Assert.AreEqual(String.Format("Status - {0}.2.4", TestGlobals.current["Cat1ID"]), result2.ElementAt(2).Key);
        }

        [TestMethod]
        public void AddCategory()
        {
            // arrange
            OpenConnection();
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

    #region Member webservice

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
                String cmd = String.Format("Remove-CSUser -UserID {0}", User);
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
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            String.Format("Remove-CSUser -UserID {0}", result);

            // assert

            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        public void CreateExistingUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            User = result;

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
            Assert.AreEqual(String.Format("{0} - user NOT created. ERROR: Error creating a new user. [E662437890]", UniqueName()), result2);
        }

    }

    [TestClass]
    public class RemoveCSUserCommandTests : PSTestFixture
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
            String cmd = "Remove-CSUser -UserID 1";

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.current["DepartmentGroupID"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSUser -UserID {0}", result);
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
            OpenConnection();
            String cmd = String.Format("Get-CSUserIDByLogin -Login {0}", TestGlobals.current["UserToRetrieve"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert
            Assert.AreEqual(result, Convert.ToInt64(TestGlobals.current["UserToRetrieveID"]));
        }

    }

    #endregion

    #region Classifications webservice

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
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", result, TestGlobals.current["ClassificationIDs"]);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
            Assert.AreEqual(String.Format("{0} - classifications applied", result), result2);
        }

    }

    #endregion

    #region Records management webservice

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
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.current["RMClassificationID"]);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
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
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.current["ParentID"]);
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.current["RMClassificationID"]);
            ExecPS(cmd);
            cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", result);

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
            Assert.AreEqual(String.Format("{0} finalised", result), result2);
        }

    }

    #endregion

    #region Physical objects webservice

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
            OpenConnection();
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
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysContainer()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
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
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysBox()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysBox -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.current["ParentID"], TestGlobals.current["ItemSubType"], TestGlobals.current["HomeLocation"]);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the exception attribute
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
            OpenConnection();
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

            // assert - captured by the exception attribute
            Assert.AreEqual(String.Format("{0} - assigned to box {1}", item, box), result);
        }

    }

    #endregion

}