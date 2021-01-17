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

            string connectionString = ConfigurationManager.ConnectionStrings["Default"].ConnectionString;
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
                    Types.SelectedValuePath = "Id";
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
                string query = "SELECT * FROM fase2 WHERE fase2.Fase1Id = @Fase1Id";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                SQLiteDataAdapter sqLiteDataAdapter = new SQLiteDataAdapter(sqLiteCommand);
                using (sqLiteDataAdapter)
                {
                    sqLiteCommand.Parameters.AddWithValue("@Fase1Id", Types.SelectedValue);
                    DataTable dataTable = new DataTable();
                    sqLiteDataAdapter.Fill(dataTable);
                    Jobs.DisplayMemberPath = "Job";
                    Jobs.SelectedValuePath = "Id";
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
                string query = "SELECT * FROM fase3 WHERE fase3.Fase2Id = @Fase2Id";
                SQLiteCommand sqLiteCommand = new SQLiteCommand(query, sqlConnection);
                SQLiteDataAdapter sqLiteDataAdapter = new SQLiteDataAdapter(sqLiteCommand);
                using (sqLiteDataAdapter)
                {
                    sqLiteCommand.Parameters.AddWithValue("@Fase2Id", Jobs.SelectedValue);
                    DataTable dataTable = new DataTable();
                    sqLiteDataAdapter.Fill(dataTable);
                    Problems.DisplayMemberPath = "Problem";
                    Problems.SelectedValuePath = "Id";
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

                var headerRow = new List<string[]>(){ new string[] { "Type", "Job", "Problem", "Note", "Full Name" } };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                var data = generateListFromDatabase();
                worksheet.Cells[2, 1].LoadFromArrays(data);
                FileInfo excelFile = new FileInfo(@"C:\Users\rsamadza\excel files\test.xlsx");
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
                    sqLiteCommand.Parameters.AddWithValue("@Type", Types.SelectedItem.ToString());
                    sqLiteCommand.Parameters.AddWithValue("@Job", Jobs.SelectedItem.ToString());
                    sqLiteCommand.Parameters.AddWithValue("@Problem", Problems.SelectedItem.ToString());
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
    }
}
