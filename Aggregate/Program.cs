using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Data.Odbc;

namespace Aggregate
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var subDirectory = GetInputPaths();
            foreach (var dir in subDirectory)
            {
                ProcessFile(dir);
            }
        }

        /// <summary>
        /// Method looks into a folder from appsetting and return all the subfolders to look into
        /// </summary>
        /// <returns></returns>
        static string[] GetInputPaths()
        {

            List<string> paths;
            var dataPath = ConfigurationManager.AppSettings["datapath"];
            string[] subDirectory =
                Directory.GetDirectories(dataPath, "b*", searchOption: SearchOption.TopDirectoryOnly);
            return subDirectory;

        }

        /// <summary>
        /// Method processes the *.accdb file in the folderpath
        /// The method looks into the folder path for an accessdb file. If the file is found,
        /// makes a connection and writes the name of user tables in the console.
        /// In case of error, it writes the error message.
        /// </summary>
        /// <param name="folderPath">Path to directory of the folder</param>
        static void ProcessFile(string folderPath)
        {
            //the file pattern is *output.accdb
            var file = Directory.GetFiles(@folderPath, "*output.accdb").FirstOrDefault();

            if (File.Exists(file))
            {
                string connetionString = null;
                connetionString = @"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + file;
                OdbcConnection odbcConnection = new OdbcConnection(connetionString);
                try
                {
                    odbcConnection.Open();
                    List<string> tableNames = new List<string>();
                    var schema = odbcConnection.GetSchema("Tables");

                    foreach (System.Data.DataRow row in schema.Rows)
                    {
                        var tableName = row["TABLE_NAME"].ToString();
                        //Exclude the system tables
                        if (!tableName.StartsWith("MSys"))
                        {
                            tableNames.Add(tableName);
                        }
                    }

                    foreach (var tableName in tableNames)
                    {
                        Console.WriteLine(tableName);
                    }

                    odbcConnection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
