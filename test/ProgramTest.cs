using System;
using System.Text;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace XlsLocalizationTool
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ProgramTest
    {

        #region paths
        public const string OUPUT_DIRECTORY = "tmp\\";
        public const string SOURCE_RES_DIR = "..\\..\\..\\..\\test\\res\\";

        public const string SOURCE_RESX_DIR = SOURCE_RES_DIR + "resx\\";
        public const string SOURCE_RESX_RES1_EN_PATH = SOURCE_RESX_DIR + "en\\res1.en.resx";
        public const string SOURCE_RESX_RES1_FR_PATH = SOURCE_RESX_DIR + "fr\\res1.fr.resx";
        public const string SOURCE_RESX_RES1_JA_PATH = SOURCE_RESX_DIR + "ja\\res1.ja.resx";
        public const string SOURCE_RESX_RES2_EN_PATH = SOURCE_RESX_DIR + "en\\res2.en.resx";
        public const string SOURCE_RESX_RES2_FR_PATH = SOURCE_RESX_DIR + "fr\\res2.fr.resx";
        public const string SOURCE_RESX_RES2_JA_PATH = SOURCE_RESX_DIR + "ja\\res2.ja.resx";

        public const string SOURCE_PROPERTIES_DIR = SOURCE_RES_DIR + "properties\\";
        public const string SOURCE_PROPERTIES_RES1_EN_PATH = SOURCE_PROPERTIES_DIR + "res1_en.properties";
        public const string SOURCE_PROPERTIES_RES1_FR_PATH = SOURCE_PROPERTIES_DIR + "res1_fr.properties";
        public const string SOURCE_PROPERTIES_RES1_JA_PATH = SOURCE_PROPERTIES_DIR + "res1_ja.properties";
        public const string SOURCE_PROPERTIES_RES2_EN_PATH = SOURCE_PROPERTIES_DIR + "res2_en.properties";
        public const string SOURCE_PROPERTIES_RES2_FR_PATH = SOURCE_PROPERTIES_DIR + "res2_fr.properties";
        public const string SOURCE_PROPERTIES_RES2_JA_PATH = SOURCE_PROPERTIES_DIR + "res2_ja.properties";

        public const string XLS_FILE_NAME = "excelSheet.xls";
        public const string INPUT_XLS_FILE_PATH = SOURCE_RES_DIR + "excel\\" + XLS_FILE_NAME;
        public const string OUTPUT_XLS_FILE_PATH = OUPUT_DIRECTORY + XLS_FILE_NAME;

        public const string INPUT_XLSX_FILE_PATH = SOURCE_RES_DIR + "excel\\" + XLS_FILE_NAME;
        public const string OUTPUT_XLSX_FILE_PATH = OUPUT_DIRECTORY + XLS_FILE_NAME;
        #endregion

        public ProgramTest()
        {
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
        //[TestInitialize()]
        //public void MyTestInitialize() {}

        enum ItemToDeploy
        {
            xls,
            xlsx,
            resx,
            properties
        }

        private void deploy(ItemToDeploy item)
        {
            DirectoryInfo targetDirectory = new DirectoryInfo(OUPUT_DIRECTORY);
            try
            {
                targetDirectory.Delete(true); //delete directory, subdirectories and files
            } catch (DirectoryNotFoundException) {}

            targetDirectory.Create();

            switch (item)
            {
                case ItemToDeploy.xls:
                    deployXlsFile();
                    break;
                case ItemToDeploy.xlsx:
                    deployXlsxFile();
                    break;
                case ItemToDeploy.resx:
                case ItemToDeploy.properties:
                default:
                    throw new Exception("unsupported item");
            }
        }

        private void deployXlsFile()
        {
            new FileInfo(INPUT_XLS_FILE_PATH).CopyTo(OUTPUT_XLS_FILE_PATH);
        }

        private void deployXlsxFile()
        {
            new FileInfo(INPUT_XLSX_FILE_PATH).CopyTo(OUTPUT_XLSX_FILE_PATH);
        }

        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestXlsToProperties()
        {
            deploy(ItemToDeploy.xls);

            XlsLocalizationManager manager = new XlsLocalizationManager();

            manager.XlsToUTF8Properties(new FileInfo(OUTPUT_XLS_FILE_PATH).FullName, "en");

            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1.properties", SOURCE_PROPERTIES_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1_fr.properties", SOURCE_PROPERTIES_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1_ja.properties", SOURCE_PROPERTIES_RES1_JA_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2.properties", SOURCE_PROPERTIES_RES2_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2_fr.properties", SOURCE_PROPERTIES_RES2_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2_ja.properties", SOURCE_PROPERTIES_RES2_JA_PATH));
        }

        [TestMethod]
        public void TestXlsToResx()
        {
            deploy(ItemToDeploy.xls);

            XlsLocalizationManager manager = new XlsLocalizationManager();

            manager.XlsToResx(new FileInfo(OUTPUT_XLS_FILE_PATH).FullName, "en");
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "en\\res1.resx", SOURCE_RESX_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "fr\\res1.fr.resx", SOURCE_RESX_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "ja\\res1.ja.resx", SOURCE_RESX_RES1_JA_PATH));

            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "en\\res1.resx", SOURCE_RESX_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "fr\\res1.fr.resx", SOURCE_RESX_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "ja\\res1.ja.resx", SOURCE_RESX_RES1_JA_PATH));
        }

        [TestMethod]
        public void TestXlsxToProperties()
        {
            deploy(ItemToDeploy.xlsx);

            XlsLocalizationManager manager = new XlsLocalizationManager();

            manager.XlsToUTF8Properties(new FileInfo(OUTPUT_XLSX_FILE_PATH).FullName, "en");

            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1.properties", SOURCE_PROPERTIES_RES1_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1_fr.properties", SOURCE_PROPERTIES_RES1_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res1_ja.properties", SOURCE_PROPERTIES_RES1_JA_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2.properties", SOURCE_PROPERTIES_RES2_EN_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2_fr.properties", SOURCE_PROPERTIES_RES2_FR_PATH));
            Assert.IsTrue(CompareFiles(OUPUT_DIRECTORY + "res2_ja.properties", SOURCE_PROPERTIES_RES2_JA_PATH));
        }

        private bool CompareFiles(string file1, string file2)
        {
            FileInfo fileInfo1 = new FileInfo(file1);
            FileInfo fileInfo2 = new FileInfo(file2);

            if (!fileInfo1.Exists) throw new Exception("file1 does not exist");
            if (!fileInfo2.Exists) throw new Exception("file2 does not exist");

            if (fileInfo1.Length != fileInfo2.Length)
                return false;

            byte[] bytesFile1 = File.ReadAllBytes(file1);
            byte[] bytesFile2 = File.ReadAllBytes(file2);

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