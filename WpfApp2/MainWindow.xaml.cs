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
            PlateTypes.ItemsSource = new List<String>() { "U1", "U2", "M1", "M2", "L1", "L2" };
            Logo.Source = new BitmapImage(new Uri(Directory.GetCurrentDirectory().Substring(0, Directory.GetCurrentDirectory().Length - 9) + "Image\\Logo.jpeg", UriKind.Absolute));
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
                MessageBox.Show("ShowTypes" + e.ToString());
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
                MessageBox.Show("showjobs" + e.ToString());
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
                MessageBox.Show("showproblems" + e.ToString());
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
                var headerRow = new List<string[]>(){ new string[] { "IDcode", "Date", "FullName", "PHName", "PlateType", "HolesNumber", "Type", "Job", "Problem", "Note" } };
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
                    string query = "INSERT INTO FinalResults (Type, Job, Problem, Note, FullName, IdCode, PHName, HolesNumber, PlateType, Date) values (@Type, @Job, @Problem, @Note, @FullName, @IdCode, @PHName, @HolesNumber, @PlateType, @Date)";
                    SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                    sqlConnection.Open();
                    sqLiteCommand.Parameters.AddWithValue("@Type", Types.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Job", Jobs.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Problem", Problems.SelectedValue);
                    sqLiteCommand.Parameters.AddWithValue("@Note", Note.Text);
                    sqLiteCommand.Parameters.AddWithValue("@FullName", getFullName());
                    sqLiteCommand.Parameters.AddWithValue("@IdCode", Id.Text);
                    sqLiteCommand.Parameters.AddWithValue("@PHName", PH.Text);
                    sqLiteCommand.Parameters.AddWithValue("@HolesNumber", Holes.Text);
                    sqLiteCommand.Parameters.AddWithValue("@PlateType", PlateTypes.Text);
                    sqLiteCommand.Parameters.AddWithValue("@Date", (DateTime.Now.Date + DateTime.Now.TimeOfDay).ToString());
                    sqLiteCommand.ExecuteScalar();
                    emptyFields();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("saveclick" + ex.ToString());
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
                string query = "SELECT IDcode, Date, FullName, PHName, PlateType, HolesNumber, Type, Job, Problem, Note FROM FinalResults";
                sqlConnection.Open();

                var cmd = new SQLiteCommand(query, sqlConnection);
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        var cell = new string[] { rdr.GetString(0), rdr.GetString(1), rdr.GetString(2), rdr.GetString(3), rdr.GetString(4), rdr.GetString(5), rdr.GetString(6), rdr.GetString(7), rdr.GetString(8), rdr.GetString(9) };
                        cellData.Add(cell);
                    }
                }
                return cellData;
            }
            catch(Exception ex) {
                MessageBox.Show("generatelistfromdatabase" + ex.Message.ToString());
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
            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 20;
            worksheet.Column(4).Width = 20;
            worksheet.Column(5).Width = 20;
            worksheet.Column(6).Width = 20;
            worksheet.Column(7).Width = 20;
            worksheet.Column(8).Width = 20;
            worksheet.Column(9).Width = 20;
            worksheet.Column(10).Width = 20;
            worksheet.Cells[borderRange].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            worksheet.Cells[borderRange].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        private void Id_IsKeyboardFocusedChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox.Text.Equals("ID Code..."))
            {
                textBox.Text = "";
            }
            else if (textBox.Text.Equals(""))
            {
                textBox.Text = "ID Code...";
            }
        }

        private void PH_IsKeyboardFocusedChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox.Text.Equals("PH Name..."))
            {
                textBox.Text = "";
            }
            else if (textBox.Text.Equals(""))
            {
                textBox.Text = "PH Name...";
            }
        }

        private void Holes_IsKeyboardFocusedChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox.Text.Equals("Holes Number..."))
            {
                textBox.Text = "";
            }
            else if (textBox.Text.Equals(""))
            {
                textBox.Text = "Holes Number...";
            }
        }
    }
}
