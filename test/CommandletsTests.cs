using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using cscommandlets;

namespace cscommandlets.tests
{

    public class TestGlobals
    {
        public static String UserName = "admin";
        public static String Password = "p@ssw0rd";
        public static String ServicesDirectory = "http://content.cgi.demo/cws/";
        public static Int64 ParentID = 2000;
        public static Int64 TemplateID = 145843;
        public static Int64 DepartmentGroupID = 1001;
    }

    public abstract class PSTestFixture
    {
        protected Runspace runspace { get; set; }

        public void CreateRunspace()
        {
            runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            String cmd = @"Import-Module .\cscommandlets.dll";
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

    }

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
        public void OpenCSConnection()
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
        public void NoConnectionAddCSProjectWorkspace()
        {
            // arrange
            String cmd = "Add-CSProjectWorkspace -Name Tester123 -ParentID 2000 -TemplateID 123456";

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void AddProjectWorkspace()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSProjectWorkspace -Name Tester123 -ParentID {0}", TestGlobals.ParentID);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();
            var result2 = ExecPS(cmd)
                .Select(o => o.BaseObject)
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
            String cmd = String.Format("Add-CSProjectWorkspace -Name Tester123 -ParentID {0} -TemplateID {1}", TestGlobals.ParentID, TestGlobals.TemplateID);

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();
            var result2 = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();
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
        public void NoConnectionAddCSFolder()
        {
            // arrange
            String cmd = @"Add-CSFolder -Name Tester123 -ParentID 2000";

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void AddFolder()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.ParentID);

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

        String UserName = "Tester123";
        Int64 User;

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void NoConnectionAddCSUser()
        {
            // arrange
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID 1001", UserName);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void CreateUser()
        {
            // arrange
            OpenConnection();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UserName, TestGlobals.DepartmentGroupID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            User = result;

            // assert
            Assert.IsInstanceOfType(result, typeof(Int64));     // created a user
        }

        [TestMethod]
        [ExpectedException(typeof(System.Management.Automation.CmdletInvocationException))]
        public void CreateExistingUser()
        {
            // arrange
            OpenConnection();
            CreateUser();
            String cmd = String.Format("Add-CSUser -Login {0} -DepartmentGroupID {1}", UserName, TestGlobals.DepartmentGroupID);

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            User = result;
            var result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Remove-CSUser -UserID {0}", result);
            ExecPS(cmd);

            // assert - captured by the attribute
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
        public void NoConnectionRemoveCSUser()
        {
            // arrange
            String cmd = "Remove-CSUser -UserID 20000000";

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<Int64>()
                .First();

            // assert - captured by the attribute
        }

        // todo test the remove user - maybe later, it's implicit in the add user

    }

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
        public void NoConnectionAddCSClassifications()
        {
            // arrange

            String cmd = "Add-CSClassifications -NodeID 20000000 -ClassificationIDs @(123456,234567)";

            // act
            var result = ExecPS(cmd)
                .Select(o => o.BaseObject)
                .Cast<String>()
                .First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void AddClassifications()
        {
            // arrange
            OpenConnection();
            String classificationIDs = "145405,145292";
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSClassifications -NodeID {0} -ClassificationIDs @({1})", result, classificationIDs);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the attribute
            Assert.AreEqual("Classifications applied", result2);
        }

    }

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

            String cmd = "Add-CSRMClassification -NodeID 20000000 -RMClassificationID 123456";

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void AddRMClassification()
        {
            // arrange
            Int64 classificationId = 144170;

            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, classificationId);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the attribute
            Assert.AreEqual("RM classification applied", result2);
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

            String cmd = "Set-CSFinaliseRecord -NodeID 20000000";

            // act
            var result = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();

            // assert - captured by the attribute
        }

        [TestMethod]
        public void FinaliseRecord()
        {
            // arrange
            Int64 classificationId = 144170;

            OpenConnection();
            String cmd = String.Format("Add-CSFolder -Name Tester123 -ParentID {0}", TestGlobals.ParentID);
            Int64 result = ExecPS(cmd).Select(o => o.BaseObject).Cast<Int64>().First();
            cmd = String.Format("Add-CSRMClassification -NodeID {0} -RMClassificationID {1}", result, classificationId);
            ExecPS(cmd);
            cmd = String.Format("Set-CSFinaliseRecord -NodeID {0}", result);

            // act
            String result2 = ExecPS(cmd).Select(o => o.BaseObject).Cast<String>().First();
            cmd = String.Format("Remove-CSNode -NodeID {0}", result);
            ExecPS(cmd);

            // assert - captured by the attribute
            Assert.AreEqual("Item finalised", result2);
        }

    }

}
