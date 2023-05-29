using System;
using System.Collections.Generic;
using System.Data.Common;
using Microsoft.AnalysisServices.Tabular;

namespace DAXCreator
{
    class Tabular
    {
        public string server_input;
        public Server server;
        public Database database;
        public Model model;
       
        // Creating a two-dimensional list for the artifacts
        public List<List<string>> artifacts = new List<List<string>>();

        public Tabular(string server_input)
        {

            this.server_input = server_input;
            Server server = new Server();
            server.Connect(server_input);
            this.server = server;
            this.database = this.server.Databases[0];
            this.model = this.database.Model;

        }
        
        public void ReadArtifacts()
        {
            // Loop through the tables
            foreach (Table table in this.model.Tables) 
            {
                // Columns
                foreach (Column column in table.Columns)
                {
                    //Console.WriteLine($"-{table.Name}.[{column.Name}]");
                    // Adding rows to the two-dimensional list
                    artifacts.Add(new List<string>() { table.Name, column.Name, "Column" });
                }
                //Measures
                foreach (Measure measure in table.Measures)
                {
                    //Console.WriteLine($"-{table.Name}.[{measure.Name}]");
                    artifacts.Add(new List<string>() { table.Name, measure.Name, "Measure" });
                }
            
            }
            
        }

        public void PrintArtifacts() {
        foreach (List<string> row in artifacts)
            {
                foreach (string details in row)
                {
                    Console.Write(details + " | ");
                }
                Console.WriteLine();
            }
        }

        public void AddMeasure(string tableName, string measureName, string expression, string displayFolder="DAXimillian")
        {
            Table table = this.model.Tables.Find(tableName);
            Measure measure = new Measure()
            {
                Name = measureName,
                Expression = expression,
                DisplayFolder = displayFolder
            };

            measure.Annotations.Add(new Annotation() { Value = "Measure Created by DAXimilian" });

            if (!table.Measures.ContainsName(measureName))
            {
                table.Measures.Add(measure);
                Console.WriteLine($"Measure created in {tableName}/ {displayFolder}");
            }
            else
            {
                table.Measures[measureName].Expression = expression;
                table.Measures[measureName].DisplayFolder = displayFolder;
                Console.WriteLine($"Measure modified");
            }

            this.model.SaveChanges();
        }
        
        public void AddCalculatedColumn(string tableName, string columnName, string expression, string displayFolder = "DAXimillian")
        {
            Table table = this.model.Tables.Find(tableName);
            CalculatedColumn column = new CalculatedColumn()
            {
                Name = columnName,
                Expression = expression,
                DisplayFolder = displayFolder

            };

            column.Annotations.Add(new Annotation() { Value = "Column Created by DAXimilian" });

            if (!table.Columns.ContainsName(columnName))
            {
                table.Columns.Add(column);
                Console.WriteLine($"Column created in {tableName}/ {displayFolder}");
            }
            else
            {
                Console.WriteLine("Column already exists");
            }

            this.model.SaveChanges();
        }

     
        public void DeleteCalculatedColumns()
        {
            foreach (Table table in this.model.Tables)
            {
                // Find calculated columns
                var calculatedColumns = new List<CalculatedColumn>();
                foreach (Column column in table.Columns)
                {
                    if (column is CalculatedColumn)
                    {
                        calculatedColumns.Add((CalculatedColumn)column);
                    }
                }

                // Delete calculated columns
                foreach (CalculatedColumn calculatedColumn in calculatedColumns)
                {
                    table.Columns.Remove(calculatedColumn);
                }
            }

            this.model.SaveChanges();
        }
    }
}


