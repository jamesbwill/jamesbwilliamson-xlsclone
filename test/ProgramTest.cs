using System;
using System.Text;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace Resx2Xls
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ProgramTest
    {
        public ProgramTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion


        //relative path for DeploymentItem
        public const string TEST_RES_DIR = "..\\..\\res\\";
        public const string TEST_EXCEL_DIR = TEST_RES_DIR + "excel\\";

        //relative path for the test methods
        public const string TEST_RESX_DIR = "..\\..\\..\\test\\res\\resx\\";
        public const string TEST_RESX_RES1_EN_PATH = TEST_RESX_DIR + "en\\res1.en.resx";
        public const string TEST_RESX_RES1_FR_PATH = TEST_RESX_DIR + "fr\\res1.fr.resx";
        public const string TEST_RESX_RES1_JA_PATH = TEST_RESX_DIR + "ja\\res1.ja.resx";
        public const string TEST_RESX_RES2_EN_PATH = TEST_RESX_DIR + "en\\res2.en.resx";
        public const string TEST_RESX_RES2_FR_PATH = TEST_RESX_DIR + "fr\\res2.fr.resx";
        public const string TEST_RESX_RES2_JA_PATH = TEST_RESX_DIR + "ja\\res2.ja.resx";

        public const string TEST_PROPERTIES_DIR = "..\\..\\..\\test\\res\\properties\\";
        public const string TEST_PROPERTIES_RES1_EN_PATH   = TEST_PROPERTIES_DIR + "res1_en.properties";
        public const string TEST_PROPERTIES_RES1_FR_PATH   = TEST_PROPERTIES_DIR + "res1_fr.properties";
        public const string TEST_PROPERTIES_RES1_JA_PATH   = TEST_PROPERTIES_DIR + "res1_ja.properties";
        public const string TEST_PROPERTIES_RES2_EN_PATH   = TEST_PROPERTIES_DIR + "res2_en.properties";
        public const string TEST_PROPERTIES_RES2_FR_PATH   = TEST_PROPERTIES_DIR + "res2_fr.properties";
        public const string TEST_PROPERTIES_RES2_JA_PATH   = TEST_PROPERTIES_DIR + "res2_ja.properties";

        [TestMethod]
        [DeploymentItem(TEST_EXCEL_DIR + "excelSheet.xls")]
        public void TestXlsToProperties()
        {
            string dir = Directory.GetCurrentDirectory() + "\\";

            Resx2Xls.Program.RunCommandLine(dir + "excelSheet.xls", "en", true);
            Assert.IsTrue(CompareFiles(dir + "res1.properties",     TEST_PROPERTIES_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(dir + "res1_fr.properties",  TEST_PROPERTIES_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(dir + "res1_ja.properties",  TEST_PROPERTIES_RES1_JA_PATH));
            Assert.IsTrue(CompareFiles(dir + "res2.properties",     TEST_PROPERTIES_RES2_EN_PATH));
            Assert.IsTrue(CompareFiles(dir + "res2_fr.properties",  TEST_PROPERTIES_RES2_FR_PATH));
            Assert.IsTrue(CompareFiles(dir + "res2_ja.properties",  TEST_PROPERTIES_RES2_JA_PATH));            
        }

        [TestMethod]
        [DeploymentItem(TEST_EXCEL_DIR + "excelSheet.xls")]
        public void TestXlsToResx()
        {
            string dir = Directory.GetCurrentDirectory() + "\\";

            Resx2Xls.Program.RunCommandLine(dir + "excelSheet.xls", "en", false);
            Assert.IsTrue(CompareFiles(dir + "en\\res1.resx", TEST_RESX_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(dir + "fr\\res1.fr.resx", TEST_RESX_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(dir + "ja\\res1.ja.resx", TEST_RESX_RES1_JA_PATH));

            Assert.IsTrue(CompareFiles(dir + "en\\res1.resx", TEST_RESX_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(dir + "fr\\res1.fr.resx", TEST_RESX_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(dir + "ja\\res1.ja.resx", TEST_RESX_RES1_JA_PATH));
        }


        private bool CompareFiles(string File1, string File2)
        {
            FileInfo FI1 = new FileInfo(File1);
            FileInfo FI2 = new FileInfo(File2);

            if (FI1.Length != FI2.Length)
                return false;

            byte[] bytesFile1 = File.ReadAllBytes(File1);
            byte[] bytesFile2 = File.ReadAllBytes(File2);

            if (bytesFile1.Length != bytesFile2.Length)
                return false;

            for (int i = 0; i <= bytesFile2.Length - 1; i++)
            {
                if (bytesFile1[i] != bytesFile2[i])
                    return false;
            }
            return true;
        }
    }
}
