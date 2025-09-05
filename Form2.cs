using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelToDB
{
    public partial class Form2 : Form
    {
        private string connectionString;

        public Form2(string connStr)
        {
            InitializeComponent();
            connectionString = connStr;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            LoadDatabases();
        }

       
        private void LoadDatabases()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    var cmd = new SqlCommand("SELECT name FROM sys.databases WHERE database_id > 4", connection);
                    using (var reader = cmd.ExecuteReader())
                    {
                        cmbDatabases.Items.Clear();
                        while (reader.Read())
                        {
                            cmbDatabases.Items.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Veritabanı listesi alınamadı: " + ex.Message);
            }
        }

        private void LoadTables()
        {
            if (cmbDatabases.SelectedItem == null) return;
            string selectedDB = cmbDatabases.SelectedItem.ToString();
            lstTables.Items.Clear();
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = selectedDB };

            try
            {
                using (var connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    var query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_SCHEMA = 'dbo' ORDER BY TABLE_NAME";
                    using (var cmd = new SqlCommand(query, connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            lstTables.Items.Add(reader.GetString(0));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Tablo listesi çekme hatası: " + ex.Message);
            }
        }

        private void cmbDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgvTableData.DataSource = null;
            LoadTables();
        }

        private void lstTables_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstTables.SelectedItem == null || cmbDatabases.SelectedItem == null) return;
            string selectedDB = cmbDatabases.SelectedItem.ToString();
            string selectedTable = lstTables.SelectedItem.ToString();
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = selectedDB };

            try
            {
                using (var connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();
                    var query = $"SELECT TOP 100 * FROM [{selectedTable}]";
                    using (var adapter = new SqlDataAdapter(query, connection))
                    {
                        var dt = new DataTable();
                        adapter.Fill(dt);
                        dgvTableData.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Tablo verisi alınamadı: " + ex.Message);
            }
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            if (cmbDatabases.SelectedItem == null || lstTables.SelectedItem == null)
            {
                MessageBox.Show("Lütfen önce bir veritabanı ve tablo seçin.");
                return;
            }

            using (var ofd = new OpenFileDialog { Filter = "Excel Dosyaları|*.xlsx;*.xls", Title = "Excel Dosyası Seç" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    ImportExcelToDatabase(ofd.FileName);
                }
            }
        }
        

        private void ImportExcelToDatabase(string filePath)
        {
            var stopwatch = Stopwatch.StartNew();
            string selectedDB = cmbDatabases.SelectedItem.ToString();
            string realTable = lstTables.SelectedItem.ToString();
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = selectedDB };
            string stagingTable = $"##Staging_{Guid.NewGuid():N}";
            DataTable dt = new DataTable();
            List<string> errorMessages = new List<string>();

            using (SqlConnection conn = new SqlConnection(builder.ConnectionString))
            {
                conn.Open();
                try
                {
                    // Excel okuma ve satır satır kontrol
                    using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook = Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ? (IWorkbook)new XSSFWorkbook(file) : new HSSFWorkbook(file);
                        ISheet sheet = workbook.GetSheetAt(0);
                        IRow headerRow = sheet.GetRow(sheet.FirstRowNum);
                        for (int j = 0; j < headerRow.LastCellNum; j++) dt.Columns.Add(headerRow.GetCell(j)?.ToString().Trim() ?? $"Sütun{j}");
                        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null || row.Cells.All(c => c == null || c.CellType == CellType.Blank)) continue;
                            DataRow dr = dt.NewRow();
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                var cellValue = row.GetCell(j)?.ToString();
                                dr[j] = string.IsNullOrWhiteSpace(cellValue) ? (object)DBNull.Value : cellValue;
                            }
                            dt.Rows.Add(dr);
                        }
                    }

                    if (dt.Rows.Count == 0) { MessageBox.Show("Excel dosyasında işlenecek veri bulunamadı."); return;}

                    // Geçici tablo oluşturma
                    string createScript = $"SELECT TOP 0 * INTO {stagingTable} FROM [{realTable}]";
                    using (SqlCommand createCmd = new SqlCommand(createScript, conn)) createCmd.ExecuteNonQuery();
                    int excelRowNumber = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        excelRowNumber++;
                        try
                        {
                            var columnNames = string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"));
                            var parameterNames = string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => $"@{c.ColumnName.Replace(" ", "").Replace("-", "")}"));
                            string insertSql = $"INSERT INTO {stagingTable} ({columnNames}) VALUES ({parameterNames})";
                            using (SqlCommand insertCmd = new SqlCommand(insertSql, conn))
                            {
                                foreach (DataColumn col in dt.Columns)
                                {
                                    insertCmd.Parameters.AddWithValue($"@{col.ColumnName.Replace(" ", "").Replace("-", "")}", dr[col]);
                                }
                                insertCmd.ExecuteNonQuery();
                            }
                        }
                        catch (SqlException ex)
                        {
                            errorMessages.Add($"Excel Satır {excelRowNumber}: {ex.Message}");
                        }
                    }

                    // Hata kontrolü
                    if (errorMessages.Any())
                    {
                        stopwatch.Stop();

                        var summaryMessage = new StringBuilder();
                        summaryMessage.AppendLine($"İşlem başarısız. {errorMessages.Count} adet hata bulundu.");
                        summaryMessage.AppendLine("Veritabanına hiçbir veri kaydedilmedi.");
                        summaryMessage.AppendLine($"\nGeçen süre: {stopwatch.Elapsed:m\\:ss\\.fff}");
                        MessageBox.Show(summaryMessage.ToString(), "İşlem İptal Edildi", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        
                        Form3 errorForm = new Form3(errorMessages);
                        errorForm.ShowDialog();
                    }
                    else
                    {
                        // Başarılı durumda veri kaydı
                        var insertColumns = string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"));
                        using (SqlCommand finalInsertCmd = new SqlCommand($"INSERT INTO [{realTable}] ({insertColumns}) SELECT {insertColumns} FROM {stagingTable}", conn))
                        {
                            int insertedCount = finalInsertCmd.ExecuteNonQuery();
                            stopwatch.Stop();
                            string successMessage = $"İşlem tamamlandı. {insertedCount} satır başarıyla eklendi.\n\nGeçen süre: {stopwatch.Elapsed:m\\:ss\\.fff}";
                            MessageBox.Show(successMessage, "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    stopwatch.Stop();
                    MessageBox.Show($"Genel bir hata oluştu: {ex.Message}\n\nGeçen süre: {stopwatch.Elapsed:m\\:ss\\.fff}", "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    using (SqlCommand dropCmd = new SqlCommand($"IF OBJECT_ID('tempdb..{stagingTable}') IS NOT NULL DROP TABLE {stagingTable}", conn))
                        dropCmd.ExecuteNonQuery();

                    lstTables_SelectedIndexChanged(null, null);
                }
            }
        }

        
        
        private void btnDownloadTemplate_Click(object sender, EventArgs e)
        {
            if (cmbDatabases.SelectedItem == null || lstTables.SelectedItem == null)
            {
                MessageBox.Show("Lütfen önce bir veritabanı ve tablo seçin.");
                return;
            }
            string selectedDB = cmbDatabases.SelectedItem.ToString();
            string selectedTable = lstTables.SelectedItem.ToString();
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = selectedDB };
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Dosyası|*.xlsx";
                sfd.Title = "Excel Şablonunu Kaydet";
                sfd.FileName = $"{selectedTable}_Sablon.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var columnNames = new List<string>();
                        using (var connection = new SqlConnection(builder.ConnectionString))
                        {
                            connection.Open();
                            var query = $"SELECT TOP 0 * FROM [{selectedTable}]";
                            using (var cmd = new SqlCommand(query, connection))
                            using (var reader = cmd.ExecuteReader())
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    columnNames.Add(reader.GetName(i));
                                }
                            }
                        }
                        IWorkbook workbook = new XSSFWorkbook();
                        ISheet sheet = workbook.CreateSheet(selectedTable);
                        IRow headerRow = sheet.CreateRow(0);
                        for (int i = 0; i < columnNames.Count; i++)
                        {
                            headerRow.CreateCell(i).SetCellValue(columnNames[i]);
                        }
                        using (var fileStream = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(fileStream);
                        }
                        MessageBox.Show("Excel şablonu başarıyla oluşturuldu!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Şablon oluşturulurken bir hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            if (cmbDatabases.SelectedItem == null || lstTables.SelectedItem == null)
            {
                MessageBox.Show("Lütfen önce bir veritabanı ve tablo seçin.");
                return;
            }
            string selectedDB = cmbDatabases.SelectedItem.ToString();
            string selectedTable = lstTables.SelectedItem.ToString();
            var builder = new SqlConnectionStringBuilder(connectionString) { InitialCatalog = selectedDB };
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Dosyası|*.xlsx";
                sfd.Title = "Verileri Excel'e Aktar";
                sfd.FileName = $"{selectedTable}_Veri.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        DataTable dt = new DataTable();
                        using (var connection = new SqlConnection(builder.ConnectionString))
                        {
                            var query = $"SELECT * FROM [{selectedTable}]";
                            using (var adapter = new SqlDataAdapter(query, connection))
                            {
                                adapter.Fill(dt);
                            }
                        }
                        IWorkbook workbook = new XSSFWorkbook();
                        ISheet sheet = workbook.CreateSheet(selectedTable);
                        IRow headerRow = sheet.CreateRow(0);
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            headerRow.CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                        }
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            IRow dataRow = sheet.CreateRow(i + 1);
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                object cellValue = dt.Rows[i][j];
                                dataRow.CreateCell(j).SetCellValue(cellValue == DBNull.Value ? "" : cellValue.ToString());
                            }
                        }
                        using (var fileStream = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(fileStream);
                        }
                        MessageBox.Show("Veriler başarıyla Excel'e aktarıldı!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Veriler aktarılırken bir hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        
    }
}