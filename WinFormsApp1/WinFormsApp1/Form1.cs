using System;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        // SQL Server connection string
        string connectionString = @"Server=DESKTOP-2CTNAI7\SQLEXPRESS;Database=db1;Trusted_Connection=True;";

        public Form1()
        {
            InitializeComponent();
        }

        // Import Excel to DataGridView
        private void Button1_Click(object sender, EventArgs e)
        {
            string filepath = @"D:\excel\Book1.xlsx"; // Update path as needed
            Excel.Application excelApp = new();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filepath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            dataGridView1.Rows.Clear(); // Clear previous data
            dataGridView1.ColumnCount = range.Columns.Count;

            for (int row = 1; row <= range.Rows.Count; row++)
            {
                DataGridViewRow dgvRow = new DataGridViewRow();

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    string? cellValue = ((Excel.Range)range.Cells[row, col]).Value?.ToString();

                    if (row == 1)
                    {
                        dataGridView1.Columns[col - 1].HeaderText = cellValue;
                    }
                    else
                    {
                        if (dgvRow.Cells.Count < range.Columns.Count)
                            dgvRow.Cells.Add(new DataGridViewTextBoxCell());

                        dgvRow.Cells[col - 1].Value = cellValue;
                    }
                }

                if (row > 1)
                    dataGridView1.Rows.Add(dgvRow);
            }

            workbook.Close(false);
            excelApp.Quit();

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
        }

        private void SaveDataToDatabase()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string name = row.Cells[0].Value?.ToString();
                        string regNo = row.Cells[1].Value?.ToString();
                        string department = row.Cells[2].Value?.ToString();

                        // ✅ Check if RegNo already exists
                        string checkQuery = "SELECT COUNT(*) FROM excel WHERE RegNo = @RegNo";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection))
                        {
                            checkCmd.Parameters.AddWithValue("@RegNo", regNo ?? "");
                            int exists = (int)checkCmd.ExecuteScalar();

                            if (exists > 0)
                            {
                                MessageBox.Show($"RegNo {regNo} is already registered.", "Duplicate Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                continue;
                            }
                        }

                        // ✅ Insert the new record
                        string insertQuery = "INSERT INTO excel (Name, RegNo, Department) VALUES (@Name, @RegNo, @Department)";
                        using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection))
                        {
                            insertCmd.Parameters.AddWithValue("@Name", name ?? "");
                            insertCmd.Parameters.AddWithValue("@RegNo", regNo ?? "");
                            insertCmd.Parameters.AddWithValue("@Department", department ?? "");

                            insertCmd.ExecuteNonQuery();
                        }
                    }
                }

                MessageBox.Show("✅ Data inserted successfully.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Save data from DataGridView to SQL Server
        private void SaveDataToDatabasefromexcel()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string name = row.Cells[0].Value?.ToString();
                        string regNo = row.Cells[1].Value?.ToString();
                        string department = row.Cells[2].Value?.ToString();

                        string query = "INSERT INTO excel (Name, RegNo, Department) VALUES (@Name, @RegNo, @Department)";

                        using (SqlCommand cmd = new SqlCommand(query, connection))
                        {
                            cmd.Parameters.AddWithValue("@Name", name ?? "");
                            cmd.Parameters.AddWithValue("@RegNo", regNo ?? "");  
                            cmd.Parameters.AddWithValue("@Department", department ?? "");

                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                MessageBox.Show("✅ Data inserted into 'excel' table successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Save Button Click
        private void Button2_Click(object sender, EventArgs e)
        {
            SaveDataToDatabase();
        }
    }
}
