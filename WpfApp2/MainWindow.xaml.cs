using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SQLite;
using System.Configuration;
using System.Data;
using OfficeOpenXml;
using System.IO;

namespace WpfApp2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SQLiteConnection sqlConnection;

        public MainWindow()
        {
            InitializeComponent();
            Directory.CreateDirectory(Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnoProbe");
            string connectionString = "Data Source="+ Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "TechnoProbe.db;Version=3;New=False;Compress=True;";
            sqlConnection =new SQLiteConnection(connectionString);
            showTypes();
        }

        public void showTypes() {
            try
            {
                string query = "SELECT * FROM fase1";
                SQLiteDataAdapter sqLiteDataAdapter = new SQLiteDataAdapter(query, sqlConnection);
                using (sqLiteDataAdapter)
                {
                    DataTable dataTable = new DataTable();
                    sqLiteDataAdapter.Fill(dataTable);
                    Types.DisplayMemberPath = "Type";
                    Types.SelectedValuePath = "Type";
                    Types.ItemsSource = dataTable.DefaultView;
                }
            }
            catch (Exception e) {
                MessageBox.Show(e.ToString());
            }
        }

        public void showJobs() {
            try
            {
                string query = "SELECT * FROM fase2 WHERE fase2.Type = @Type";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                SQLiteDataAdapter sqLiteDataAdapter = new SQLiteDataAdapter(sqLiteCommand);
                using (sqLiteDataAdapter)
                {
                    sqLiteCommand.Parameters.AddWithValue("@Type", Types.SelectedValue);
                    DataTable dataTable = new DataTable();
                    sqLiteDataAdapter.Fill(dataTable);
                    Jobs.DisplayMemberPath = "Job";
                    Jobs.SelectedValuePath = "Job";
                    Jobs.ItemsSource = dataTable.DefaultView;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public void showProblems() {
            try
            {
                string query = "SELECT * FROM fase3 WHERE fase3.Job = @Job";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                SQLiteDataAdapter sqLiteDataAdapter = new SQLiteDataAdapter(sqLiteCommand);
                using (sqLiteDataAdapter)
                {
                    sqLiteCommand.Parameters.AddWithValue("@Job", Jobs.SelectedValue);
                    DataTable dataTable = new DataTable();
                    sqLiteDataAdapter.Fill(dataTable);
                    Problems.DisplayMemberPath = "Problem";
                    Problems.SelectedValuePath = "Problem";
                    Problems.ItemsSource = dataTable.DefaultView;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            generateListFromDatabase();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");
                var data = generateListFromDatabase();
                var headerRow = new List<string[]>(){ new string[] { "Type", "Job", "Problem", "Note", "Full Name" } };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                string borderRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + (data.Count + 1).ToString();
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                setExcelStyle(worksheet, headerRange, borderRange);
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[2, 1].LoadFromArrays(data);
                var dir = Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()) + "TechnoProbe";
                FileInfo excelFile = new FileInfo(dir +"\\TechnoProb.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (checkFields())
            {
                try
                {
                    string query = "INSERT INTO FinalResults (Type, Job, Problem, Note, FullName) values (@Type, @Job, @Problem, @Note, @FullName)";
                    SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                    sqlConnection.Open();
                    sqLiteCommand.Parameters.AddWithValue("@Type", Types.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Job", Jobs.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Problem", Problems.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Note", Note.Text);
                    sqLiteCommand.Parameters.AddWithValue("@FullName", getFullName());
                    sqLiteCommand.ExecuteScalar();
                    emptyFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    sqlConnection.Close();
                }
            } else
            {
                MessageBox.Show("Please Fill all Fields");
            }
        }

        private void Types_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            showJobs();
        }

        private void Jobs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            showProblems();
        }

        private void Note_IsKeyboardFocusedChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox.Text.Equals("Note...")) {
                textBox.Text = "";
            }
            else if (textBox.Text.Equals(""))
            {
                textBox.Text = "Note...";
            }
        }

        public bool checkFields() {
            if (Name.Text.Equals("") || Surname.Text.Equals("") || Note.Text.Equals("Note...") ||
                Note.Text.Equals("") || Jobs.SelectedIndex == -1 || Problems.SelectedIndex == -1 ||
                Types.SelectedIndex == -1) {
                return false;
            }
            return true;
        }

        public String getFullName() {
            return Name.Text + " " + Surname.Text;
        }

        public void emptyFields() {
            Note.Text = "Note...";
            Name.Text = "";
            Surname.Text = "";
            Jobs.ItemsSource = null;
            Problems.ItemsSource = null;
            Types.SelectedItem = null;
        }

        public List<string []> generateListFromDatabase() {
            var cellData = new List<string []>();
            try
            {
                string query = "SELECT Type, Job, Problem, Note, FullName FROM FinalResults";
                sqlConnection.Open();

                var cmd = new SQLiteCommand(query, sqlConnection);
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        var cell = new string[] { rdr.GetString(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3), rdr.GetString(4) };
                        cellData.Add(cell);
                    }
                }
                return cellData;
            }
            catch(Exception ex) {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
            return cellData;
        }

        public void setExcelStyle(OfficeOpenXml.ExcelWorksheet worksheet, String headerRange, String borderRange) {
            worksheet.Cells[headerRange].Style.Font.Bold = true;
            worksheet.Cells[headerRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[headerRange].Style.Fill.BackgroundColor.SetColor(0, 150, 150, 150);
            worksheet.Cells[headerRange].AutoFitColumns();
            worksheet.Cells[borderRange].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }
    }
}
