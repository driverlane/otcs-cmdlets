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

        // the folder where the test objects will be created
        public static Int64 ParentID = 153443;
        //public static Int64 ParentID = 123435;

        #region Encryption

        #endregion

        #region Connection

        // Open-CSConnection
        public static String UserName = "admin";
        public static String Password = "p@ssw0rd";
        public static String ServicesDirectory = "http://content.cgi.demo/cws/";
        //public static String ServicesDirectory = "http://content.cgi.demo/les-services/";

        #endregion

        #region Document management webservice

        // Add-CSProjectWorkspace
        public static Int64 TemplateID = 145843;
        //public static Int64 TemplateID = 123094;

        // cats and atts testing
        public static Int64 Cat1ID = 140266;
        public static String Cat1Name = "uiChanges testing";
        public static Int64 Cat1Version = 8;
        public static Int64 Cat2ID = 73165;
        public static String Cat2Name = "Drawing";
        public static Int64 Cat2Version = 4;

        #endregion

        #region Member webservice

        public static Int64 DepartmentGroupID = 1001;

        // Get-CSUserIDByLogin
        public static String UserToRetrieve = "admin";
        public static Int64 UserToRetrieveID = 1000;

        #endregion

        #region Classifications webservice

        public static String ClassificationIDs = "145405,145292";
        //public static String ClassificationIDs = "121557,123433";

        #endregion

        #region Records management webservice

        // Add-CSRMClassification
        public static Int64 RMClassificationID = 144170;
        //public static Int64 RMClassificationID = 114514;

        #endregion

        #region Physical objects webservice

        public static Int64 ItemSubType = 129705;
        //public static Int64 ItemSubType = 123319;
        public static Int64 PartSubType = 130586;
        //public static Int64 PartSubType = 123209;
        public static Int64 BoxSubType = 129707;
        //public static Int64 BoxSubType = 122657;
        public static String HomeLocation = "Compactus Level One Building 3";

        #endregion

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
            String cmd = String.Format("Open-CSConnection -Username {0} -Password {1} -ServicesDirectory {2}", TestGlobals.UserName, TestGlobals.Password, TestGlobals.ServicesDirectory);
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

        // todo test the encryption and logging in with encrypted password
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
            String cmd = String.Format("Open-CSConnection -Username {0} -Password {1} -ServicesDirectory {2}", TestGlobals.UserName, TestGlobals.Password, TestGlobals.ServicesDirectory);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<String>()
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
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddProjectWorkspace()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();
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
            String cmd = String.Format("Add-CSProjectWorkspace -Name {0} -ParentID {1} -TemplateID {2}", UniqueName(), TestGlobals.ParentID, TestGlobals.TemplateID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
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
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddFolder()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a folder
            Assert.AreEqual(result, result2);                   // the second run returned the ID from the first run

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
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveNode()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

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
            String cmd = String.Format("Get-CSCategories -NodeID {0}", 123456);

            // act
            String result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddCategory()
        {
            // arrange
            String cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {0}", 123456);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void GetCategories()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);
            var item = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

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
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);
            var item = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Get-CSAttributeValues -NodeID {0} -CategoryID {1}", item, TestGlobals.Cat1ID);
            Dictionary<String, List<Object>> result1 = (Dictionary<String, List<Object>>)ExecPS(cmd).Select(o => o.BaseObject).First();

            // get the attributes on the populated folder
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.Cat1ID);
            ExecPS(cmd);
            cmd = String.Format("Get-CSAttributeValues -NodeID {0} -CategoryID {1}", item, TestGlobals.Cat1ID);
            Dictionary<String, List<Object>> result2 = (Dictionary<String, List<Object>>)ExecPS(cmd).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(0, result1.Count);
            Assert.AreEqual(20, result2.Count);
        }

        [TestMethod]
        public void AddCategory()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);
            var item = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.Cat1ID);
            ExecPS(cmd);
            String cmd1 = String.Format("Get-CSCategories -NodeID {0}", item); 
            String cmd2 = String.Format("Get-CSCategories -NodeID {0} -ShowKey", item);

            // act
            List<String> result1 = (List<String>)ExecPS(cmd1).Select(o => o.BaseObject).First();
            List<String> result2 = (List<String>)ExecPS(cmd2).Select(o => o.BaseObject).First();
            cmd = String.Format("Add-CSCategory -NodeID {0} -CategoryID {1}", item, TestGlobals.Cat2ID);
            ExecPS(cmd);
            List<String> result3 = (List<String>)ExecPS(cmd1).Select(o => o.BaseObject).First();
            List<String> result4 = (List<String>)ExecPS(cmd2).Select(o => o.BaseObject).First();

            // clean up
            cmd = String.Format("Remove-CSNode -NodeID {0}", item);
            ExecPS(cmd);

            // assert
            Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.Cat1Name), result1.First());
            Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.Cat1Name, TestGlobals.Cat1ID, TestGlobals.Cat1Version), result2.First());
            Assert.AreEqual(2, result3.Count);
            foreach (String cat in result3)
            {
                if (cat.Contains(TestGlobals.Cat1Name))
                {
                    Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.Cat1Name), cat);
                }
                else
                {
                    Assert.AreEqual(String.Format("{0} - {1}", item, TestGlobals.Cat2Name), cat);
                }

            }
            Assert.AreEqual(2, result4.Count);
            foreach (String cat in result4)
            {
                if (cat.Contains(TestGlobals.Cat1Name))
                {
                    Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.Cat1Name, TestGlobals.Cat1ID, TestGlobals.Cat1Version), cat);
                }
                else
                {
                    Assert.AreEqual(String.Format("{0} - {1} - {2}.{3}", item, TestGlobals.Cat2Name, TestGlobals.Cat2ID, TestGlobals.Cat2Version), cat);
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
            if (User > 0)
            {
                String cmd = String.Format("Remove-CSUser -UserID {0}", User);
                ExecPS(cmd);
            }
            CloseRunspace();
            
        }

        Int64 User;

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddUser()
        {
            // arrange
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.DepartmentGroupID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void CreateUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.DepartmentGroupID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            String.Format("Remove-CSUser -UserID {0}", result);

            // assert

            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        public void CreateExistingUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.DepartmentGroupID);
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            User = result;

            // act
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

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
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void RemoveUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UniqueName(), TestGlobals.DepartmentGroupID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Remove-CSUser -UserID {0}", result);
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

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
            String cmd = String.Format("Get-CSUserIDByLogin -Login {0}", TestGlobals.UserToRetrieve);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void GetUserID()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Get-CSUserIDByLogin -Login {0}", TestGlobals.UserToRetrieve);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert
            Assert.AreEqual(result, TestGlobals.UserToRetrieveID);
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
            String cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", TestGlobals.ParentID, TestGlobals.ClassificationIDs);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddClassifications()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", result, TestGlobals.ClassificationIDs);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
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
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", TestGlobals.ParentID, TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddRMClassification()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name {0} -ParentID {1}", UniqueName(), TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.RMClassificationID);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
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
            String cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void FinaliseRecord()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, TestGlobals.RMClassificationID);
            ExecPS(cmd);
            cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", result);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
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
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysItem()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
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
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysContainer()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysContainer -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
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
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AddPhysBox()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysBox -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
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
            String cmd = "Set-CSPhysObjToBox -ItemID 20000000 -BoxID 123456";

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the exception attribute
        }

        [TestMethod]
        public void AssignToBox()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSPhysItem -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);
            var item = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSPhysBox -Name {0} -ParentID {1} -PhysicalItemSubType {2} -HomeLocation \"{3}\"", UniqueName(), TestGlobals.ParentID, TestGlobals.ItemSubType, TestGlobals.HomeLocation);
            var box = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Set-CSPhysObjToBox -ItemID {0} -BoxID {1}", item, box);

            // act
            String result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
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
