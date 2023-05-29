using System;
using System.Data;
using DAXCreator;

namespace DAXimilian
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string ssas_server = ServerFinder.FindServer();
            Tabular tabular = new Tabular(ssas_server);
            Console.WriteLine($"Connected to {ssas_server}");
            
            //Print and read the artifacts:
            //tabular.ReadArtifacts();
            //tabular.PrintArtifacts();

            Console.WriteLine("<<<Menu>>>");
            Console.WriteLine("1-Add Columns and measures manually");
            Console.WriteLine("2-Add Columns and measures using an Excel file");
            Console.WriteLine(" Note: The excel file must contain the columns:");
            Console.WriteLine(" [Name, Type, Table, Expression, Folder]");
            Console.WriteLine();
            Console.Write("Choose an option (1,2): ");
            string option = Console.ReadLine();
            option= option.Trim().ToLower();

            if (option=="1") 
            {

                AddArtifacts(tabular);
            }

             else if (option=="2")
            {
                AddWithExcel(tabular);
            }

            else 
            { 
                Console.WriteLine("Invalid option. Press any key to exit.");
                Console.ReadLine();

            
            }
        }

        public static void AddWithExcel(Tabular tabular) 
        {
            Console.Write("Complete path to .xlsx file: ");
            string path = Console.ReadLine();

         
            DataTable dt = ExcelREader.ExcelToDataTable(path);
     
            foreach (DataRow row in dt.Rows)
            {
                string name = row["Name"].ToString();
                string type = row["Type"].ToString().ToLower().Trim();
                string table = row["Table"].ToString();
                string expression = row["Expression"].ToString();
                string displayFolder = row["Folder"].ToString();

                if (type=="measure")
                {
                    Console.WriteLine($"Creating Name: {name}, Type: {type}");
                    tabular.AddMeasure(table, name, expression, displayFolder);
                }

                if (type == "column")
                {
                    Console.WriteLine($"Creating Name: {name}, Type: {type}");
                    tabular.AddCalculatedColumn(table, name, expression, displayFolder);
                }


            }
            Console.ReadLine();
        }

        public static void AddArtifacts(Tabular tabular)
        {
            Console.WriteLine();
            Console.WriteLine("<<<Add Artifacts>>>");
            Console.Write("-Add a measure (m), caclulated column (c) or skip (n)?:"); string firstAns = Console.ReadLine();
            firstAns = firstAns.ToLower().Trim();

            if (firstAns == "m")
            {
                bool runAgain = true;

                while (runAgain)
                {

                    Console.Write("--Measure name: ");
                    string name = Console.ReadLine();
                    Console.Write("--Measure table: ");
                    string table = Console.ReadLine();
                    Console.Write("--Folder: ");
                    string displayFolder = Console.ReadLine();
                    Console.Write("--Expression: ");
                    string expression = Console.ReadLine();


                    tabular.AddMeasure(table, name, expression, displayFolder);
                    Console.Write("Add another measure (y/n)?");
                    string runAgainText = Console.ReadLine();
                    runAgainText.ToLower().Trim();

                    if (runAgainText != "y")
                    {
                        runAgain = false;
                    }

                }


            }

            else if (firstAns == "c")
            {
                bool runAgain = true;

                while (runAgain)
                {

                    Console.Write("--Column name: ");
                    string name = Console.ReadLine();
                    Console.Write("--Column table: ");
                    string table = Console.ReadLine();
                    Console.Write("--Folder: ");
                    string displayFolder = Console.ReadLine();
                    Console.Write("--Expression: ");
                    string expression = Console.ReadLine();


                    tabular.AddCalculatedColumn(table, name, expression, displayFolder);
                    Console.Write("Add another column (y/n)?");
                    string runAgainText = Console.ReadLine();
                    runAgainText.ToLower().Trim();

                    if (runAgainText != "y")
                    {
                        runAgain = false;
                    }

                }


            }

            else
            {
                Console.Write("Press any button to leave:");
                Console.ReadLine();
            }


        }
    }
}

