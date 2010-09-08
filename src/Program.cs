using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;

namespace Resx2Xls
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] args)
        {

            if (args.Length > 0)
            {
                String infile = args[0];
                String defaultLang = String.Empty;
                Boolean generateUtf8PropertiesFile = false;

                try
                {

                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].Equals("-file", StringComparison.InvariantCultureIgnoreCase))
                        infile = args[++i];
                    else if (args[i].Equals("-dl", StringComparison.InvariantCultureIgnoreCase))
                        defaultLang = args[++i];
                    else if (args[i].Equals("-pp", StringComparison.InvariantCultureIgnoreCase))
                        generateUtf8PropertiesFile = args[++i].ToLower().Equals("yes");
                    else
                        throw new Exception();
                }
                } 
                catch
                {
                    Usage();
                    Environment.Exit(1);
                }
                
                RunCommandLine(infile, defaultLang, generateUtf8PropertiesFile);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Resx2XlsForm());
            }
        }

        static void Usage()
        {
            Console.WriteLine("Usage: XlsLocalizationTool.exe -file [filename] -dl [default language] -pp [yes/no]");
            Console.WriteLine("\tfile\tThe target file of the operation");
            Console.WriteLine("\tdl\tThe language to treat as default");
            Console.WriteLine("\tpp\tChoose yes to generate properties file, no to generate resx file. Default is no.");
        }

        static void RunCommandLine(string infile, string defaultLang, Boolean generateUtf8PropertiesFile)
        {
            Resx2XlsForm form = new Resx2XlsForm();
            using (form)
            {

                FileInfo file = new FileInfo(infile);

                try
                {
                    if (file.Exists)
                    {

                        if (file.Extension.Equals(".xls", StringComparison.InvariantCultureIgnoreCase) && generateUtf8PropertiesFile)
                        {
                            form.XlsToUTF8Properties(infile, defaultLang);
                        }
                        else if (file.Extension.Equals(".xls", StringComparison.InvariantCultureIgnoreCase) && !generateUtf8PropertiesFile) 
                        {
                            form.XlsToResx(infile, defaultLang);
                        }
                        else
                        {
                            Console.WriteLine("Command not recognized");
                            throw new Exception("Unknown filetype");
                        }
                    }
                    else
                    {
                        throw new Exception(String.Format("{0} doesn't exist", infile));
                    }
                    Console.WriteLine("Done.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Environment.Exit(1);
                }
            }
        }

    }
}
