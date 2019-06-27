using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System.Data.Odbc;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Diagnostics;

namespace CombineTables
{
    class Program
    {

        class Input
        {
            public String CombinedOutputFileName { get; set; }
            public String[] TableNames { get; set; }
            public String[] DataBases { get; set; }
        }

        private static Input input = null;
        private static Application app = null;
        private static List<String> tempFiles = new List<string>();

        static void CleanUpTempFiles()
        {
            foreach (var tempFile in tempFiles)
            {
                try
                {
                    File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error Deleting File {tempFile} \n   {ex.Message} ");
                }
                
            }

            app.Quit();
        }
        static void Main(string[] args)
        {
            //Tokenize args
            //TODO Create Help Structure
            //for (int i = 0; i < args.Length; i++)
            //{
            //    if (args[i].Equals("h", StringComparison.CurrentCultureIgnoreCase) ||
            //        args[i].Equals("help", StringComparison.CurrentCultureIgnoreCase))
            //    {



            //    }

            //}

            Run();
            return;
            using (StreamReader r = new StreamReader("readme.txt"))
            {
                Console.WriteLine(r.ReadToEnd());
            }

            
            
            Console.WriteLine("Press <Enter> to continue... or any other key to exit");
            var command = Console.ReadKey();
            if (command.Key == ConsoleKey.Enter)
            {
                Console.Clear();
                Run();
            }

        }

        private static void Run()
        {
            Stopwatch sw = new Stopwatch();
            app = new Application();
            Console.WriteLine("Begin program to combine tables...");
            input = new Input();
            LoadJson();

            sw.Start();
            List<Task> tasks = new List<Task>();
            Task cleanup = Task.Run(() => CleanOutputDb());
            tasks.Add(cleanup);

            cleanup.Wait();

            ConcurrentQueue<String> FileQueue = new ConcurrentQueue<String>();

            foreach (var db in input.DataBases)
            {
                FileQueue.Enqueue(db);
            }


            while (FileQueue.Count > 1)
            {
                MakeCombineTasks(FileQueue, tasks);
                Task.WaitAll(tasks.ToArray());
            }

            string final;
            if (FileQueue.TryDequeue(out final))
            {
                File.Copy(final, input.CombinedOutputFileName, true);
            }

            ;
            sw.Stop();
            Console.WriteLine("Elapsed={0}", sw.Elapsed);
            CleanUpTempFiles();
        }

        private static void MakeCombineTasks(ConcurrentQueue<string> FileQueue, List<Task> tasks)
        {
            while (FileQueue.Count > 1)
            {
                string first, second;
                FileQueue.TryDequeue(out first);
                FileQueue.TryDequeue(out second);
                //Task t = Task.Run(() => CombineTables(first, second, input.TableNames.ToList(), FileQueue));
                Task t = Task.Run(() => CombineTables(first, second, input.TableNames.ToList(), FileQueue));
            
                tasks.Add(t);
            }
            Task.WaitAll(tasks.ToArray());
        }


        private static void CleanOutputDb()
        {
            string query = null;
            var file = input.CombinedOutputFileName;
            if (File.Exists(file))
            {
                Console.Write($"Cleaning Up Output Db {file}");
                string connetionString = null;
                connetionString = @"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + file;
                OdbcConnection odbcConnection = new OdbcConnection(connetionString);
                try
                {
                    odbcConnection.Open();
                    List<string> tableNames = input.TableNames.ToList();
                    Console.Write(" Deleted ");
                    foreach (var tableName in tableNames)
                    {
                        query = "Delete * From " + tableName + ";";
                        OdbcCommand command = new OdbcCommand(query);
                        command.Connection = odbcConnection;
                        int result = command.ExecuteNonQuery();
                        Console.Write($"{result} {tableName} records...");
                    }
                    Console.Write("\n");
                    odbcConnection.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                //Call compact and repair after cleaning the tables.
                CompactAndRepair(file);
            }
            else
            {
                var engine = new DBEngine();
                var dbs = engine.CreateDatabase(input.CombinedOutputFileName, ";LANGID=0x0409;CP=1252;COUNTRY=0", DatabaseTypeEnum.dbVersion150);
                dbs.Close();
                dbs = null;
                Console.WriteLine($"Created {input.CombinedOutputFileName}.");
            }
            
        }

        private static void CompactAndRepair(string file)
        {
            try
            {
                
                string tempFile = Path.Combine(Path.GetDirectoryName(file),
                    Path.GetRandomFileName() + Path.GetExtension(file));
                app.CompactRepair(file, tempFile, false);
                FileInfo temp = new FileInfo(tempFile);
                temp.CopyTo(file, true);
                temp.Delete();
                //Console.WriteLine($"    Compacted {file}.");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error compacting the file. Please close the file and manually compact " + ex.Message);
            }
        }

        private static string CombineTables(string file, string file1, List<String> tableNames, ConcurrentQueue<String> FileQueue)
        {
            if (!File.Exists(file1) && !File.Exists(file))
            {
                return null;
            }

            if (!File.Exists(file1))
            {
                FileQueue.Enqueue(file);
                return file;
            }

            if (!File.Exists(file))
            {
                FileQueue.Enqueue(file1);
                return file1;
            }
            Application app = new Application();
            string tempFile = Path.Combine(Path.GetDirectoryName(file),
                Path.GetRandomFileName() + Path.GetExtension(file));
            tempFiles.Add(tempFile);
            app.CompactRepair(file, tempFile, false);
            FileQueue.Enqueue(tempFile);
            FileInfo temp = new FileInfo(tempFile);
            CopyTables(tableNames, tempFile, file1);
            return tempFile;
        }

        private static void CombineTables()
        {

            var dBs = input.DataBases;
            string query = null;

            if (File.Exists(input.CombinedOutputFileName))
            {
                //if outputdb exists start a new connection

                List<string> tableNames = input.TableNames.ToList();

                //we are goign to look inside each of the dbFile2
                foreach (var db in dBs)
                {
                    if (File.Exists(db))
                    {
                        CopyTables(tableNames, input.CombinedOutputFileName, db);
                    }
                }
            }
        }

        private static void CopyTables(List<string> tableNames, string dbFile1, string dbFile2, bool compact = true)
        {
            string connetionString = @"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + dbFile1;
            string query;
            OdbcConnection odbcConnection = new OdbcConnection(connetionString);
            odbcConnection.Open();
            try
            {
                foreach (var tableName in tableNames)
                {
                    query = "Insert into " + tableName + " Select * from " + tableName + " in '" + dbFile2 + "';";
                    OdbcCommand command = new OdbcCommand(query);
                    command.Connection = odbcConnection;
                    int result = command.ExecuteNonQuery();
                    command.Dispose();
                    Console.WriteLine($"Copied {result} {tableName} records from {dbFile2}.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                odbcConnection.Close();
                odbcConnection.Dispose();
                if (compact) CompactAndRepair(dbFile1);
            }
        }


        static void LoadJson()
        {
            string json = null;
            try
            {
                using (StreamReader r = new StreamReader("input.json"))
                {
                    json = r.ReadToEnd();
                    input = JsonConvert.DeserializeObject<Input>(json);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading input data: " + json + " \n" + ex.Message);

            }
        }
    }
}
