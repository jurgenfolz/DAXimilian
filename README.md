# DAXimilian

DAXimilian provides a console interface for managing Tabular Models in Power BI Microsoft Analysis Services. It was designed for adding columns and measures from an Excel file.

## Description

The main program provides two options:

1. Add artifacts manually to the model.
2. Add artifacts to the model using an Excel file.

The artifacts can be either measures or calculated columns. The Excel file should contain a table with the following headers:

- `Name`: Name of the artifact.
- `Type`: Type of the artifact (`measure` or `column`).
- `Table`: Name of the table where the artifact should be added.
- `Expression`: DAX expression for the artifact.
- `Folder`: Display folder for the artifact.

## How to Use

1. Run the program. It will ask you to enter the server address. Provide the address of your Analysis Services server.

2. The program will display two options:

    - `1`: Add artifacts manually.
    - `2`: Add artifacts using an Excel file.

3. If you choose option `1`, follow the prompts to add the artifacts manually.

4. If you choose option `2`, you'll be asked to provide the full path to the Excel file. After you enter the path, the program will read the data from the Excel file and add the artifacts to the model.

Here's how to use option `2`:

# Excel Reader

The `ExcelREader` class provides a static method `ExcelToDataTable` that reads data from an Excel worksheet and converts it into a `DataTable`.

## Description

The `ExcelREader` class uses the EPPlus library to read an Excel file. The `ExcelToDataTable` method reads the data from the first worksheet of the Excel file and transforms it into a `DataTable`.

## Method

- `ExcelToDataTable(string path)`: This method reads the data from the first worksheet of an Excel file given its path and converts it into a `DataTable`.


# DAXCreator

The `DAXCreator` namespace contains the `Tabular` class which is used for managing Tabular Models in Microsoft Analysis Services. The class enables you to read artifacts, add measures and calculated columns, and delete calculated columns from a Tabular model.

## Description

The `Tabular` class establishes a connection to a server, then retrieves and interacts with the first database and its model on the server.

## Methods

- `Tabular(string server_input)`: Constructor of the `Tabular` class. It takes the server address as a parameter and connects to the server. 

- `ReadArtifacts()`: Method for reading tables, columns, and measures from the model and storing them in a two-dimensional list.

- `PrintArtifacts()`: Method for printing out the artifacts stored in the two-dimensional list.

- `AddMeasure(string tableName, string measureName, string expression, string displayFolder="DAXimillian")`: Method for adding a new measure to a specific table in the model.

- `AddCalculatedColumn(string tableName, string columnName, string expression, string displayFolder = "DAXimillian")`: Method for adding a new calculated column to a specific table in the model.

- `DeleteCalculatedColumns()`: Method for deleting all calculated columns in the model.


# ServerFinder

The `ServerFinder` class in C# is used to find the server running the PowerBI service on the localhost.

## Description

This class provides a method to retrieve the localhost address of the PowerBI service by finding the `msmdsrv` process, extracting its command line, and searching for the port number in the `msmdsrv.port.txt` file in the data directory.

## Methods

- `FindServer()`: Main method of the class that finds and returns the PowerBI server address on the localhost, or `null` if the server is not found.

- `GetPowerBILocalhost()`: Helper method that finds the `msmdsrv` process, extracts its command line, and searches for the port number in the `msmdsrv.port.txt` file in the data directory.

- `GetCommandLine(Process process)`: Helper method that retrieves the command line of a given process.

## Usage


```csharp
// Create an instance of the Tabular class
Tabular tabular = new Tabular("server_address");

// Choose option 2
AddWithExcel(tabular);

// Excel reader
string path = @"C:\path\to\your\excel\file.xlsx";
DataTable dt = ExcelREader.ExcelToDataTable(path);

foreach (DataRow row in dt.Rows)
{
    foreach (var item in row.ItemArray)
    {
        Console.Write($"{item} ");
    }
    Console.WriteLine();
}

//Daxcreator
// Create an instance of the Tabular class
Tabular tabular = new Tabular("server_address");

// Read artifacts
tabular.ReadArtifacts();

// Print artifacts
tabular.PrintArtifacts();

// Add a measure
tabular.AddMeasure("Sales", "Total Sales", "SUM('Sales'[Amount])");

// Add a calculated column
tabular.AddCalculatedColumn("Sales", "Profit", "[Revenue] - [Cost]");

// Delete calculated columns
tabular.DeleteCalculatedColumns();

//Server finder
ServerFinder serverFinder = new ServerFinder();
string serverAddress = serverFinder.FindServer();

if (serverAddress != null)
{
    Console.WriteLine($"Server found at: {serverAddress}");
}
else
{
    Console.WriteLine("Server not found");
}
