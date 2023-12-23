using Microsoft.UI.Xaml;
using Syncfusion.UI.Xaml.DataGrid;
using System.Data;
using System.Data.SQLite;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using Windows.Storage.Pickers;
using Windows.Storage;
using WinRT.Interop;
using System;
using Sylvan.Data.Excel;
using System.Collections.Generic;

// ... Other using directives ...

namespace SyncFusionDataGrid
{
    public sealed partial class MainWindow : Window
    {
        private SQLiteConnection sqliteConnection;

        public MainWindow()
        {
            this.InitializeComponent();
            //InitializeDatabase();
        }

        public List<Dictionary<string, object>> ReadExcelFile(string filePath)
        {
            var records = new List<Dictionary<string, object>>();
            var opts = new ExcelDataReaderOptions
            {
                GetErrorAsNull = true,
            };
            using var edr = ExcelDataReader.Create(filePath, opts);

            while (edr.Read())
            {
                var record = new Dictionary<string, object>();
                for (int i = 0; i < edr.FieldCount; i++)
                {
                    string columnName = edr.GetName(i);
                    object value = edr.IsDBNull(i) ? null : edr.GetValue(i);
                    record[columnName] = value;
                }
                records.Add(record);
            }

            return records;
        }

        private void InitializeDatabase()
        {
            sqliteConnection = new SQLiteConnection("Data Source=MyDatabase.sqlite; Version=3;");
            sqliteConnection.Open();

            // Create table - adjust this according to your Excel structure
            string sql = "CREATE TABLE IF NOT EXISTS ExcelData (Column1 TEXT, Column2 TEXT, Column3 TEXT)";
            SQLiteCommand command = new SQLiteCommand(sql, sqliteConnection);
            command.ExecuteNonQuery();
        }

        private async void OnLoadExcelFileClick(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker();
            picker.ViewMode = PickerViewMode.Thumbnail;
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            picker.FileTypeFilter.Add(".xlsx");

            // Initialize file picker with the window handle
            IntPtr hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                /*await ImportExcelToDatabase(file);
                LoadDataFromDatabase();*/
                sfDataGrid.ItemsSource = ReadExcelFile(file.Path);
            }
        }

        private async System.Threading.Tasks.Task ImportExcelToDatabase(StorageFile file)
        {
            using (var stream = await file.OpenStreamForReadAsync())
            {
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                // Import each row into the database
                for (int row = 1; row <= sheet.LastRowNum; row++) // Assuming first row is headers
                {
                    IRow rowData = sheet.GetRow(row);
                    if (rowData != null)
                    {
                        // Adjust the INSERT statement based on your Excel and table structure
                        string sql = $"INSERT INTO ExcelData (Column1, Column2, Column3) VALUES (@col1, @col2, @col3)";
                        SQLiteCommand command = new SQLiteCommand(sql, sqliteConnection);
                        command.Parameters.AddWithValue("@col1", rowData.GetCell(0)?.ToString());
                        command.Parameters.AddWithValue("@col2", rowData.GetCell(1)?.ToString());
                        command.Parameters.AddWithValue("@col3", rowData.GetCell(2)?.ToString());
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void LoadDataFromDatabase()
        {
            string sql = "SELECT * FROM ExcelData";
            SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(sql, sqliteConnection);
            DataTable dataTable = new DataTable();
            dataAdapter.Fill(dataTable);
            sfDataGrid.ItemsSource = dataTable;
        }
    }
}