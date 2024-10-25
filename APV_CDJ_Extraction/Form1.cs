using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml; // Ensure you have this for Excel file handling

namespace APV_CDJ_Extraction
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Upload.Click += Upload_Click;
            LoadComboBoxItems();
            UploadDb.Click += UploadDb_Click;
            LoadStoreNamesIntoGridView(); // Load store names when the form initializes
            LoadTOCNamesIntoGridView(); // Load TOC names when the form initializes
        }
        private void Upload_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*";
            openFileDialog1.Title = "Select an Excel File";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;

                // Update the label to display the file name
                filename.Text = "File Name: " + Path.GetFileName(filePath);

                // Process the Excel file and load data into DataGridView
                LoadExcelData(filePath);
            }
        }

        private void LoadExcelData(string filePath)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();

                // Add all the columns from the Excel file to dataGridView1
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataGridView1.Columns.Add(worksheet.Cells[1, col].Text, worksheet.Cells[1, col].Text);
                }

                // Add the new "ExtractedStoreName" and "ExtractedTOCs" columns
                dataGridView1.Columns.Add("ExtractedStoreName", "Extracted Store Name");
                dataGridView1.Columns.Add("ExtractedTOCs", "Extracted TOCs"); // New column for TOCs

                // Load the data from the Excel sheet into dataGridView1
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataGridViewRow dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView1);

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (col == 1) // Check if it's the first column
                        {
                            string cellValue = worksheet.Cells[row, col].Text;
                            dataGridViewRow.Cells[col - 1].Value = /*ExtractRfpNumber*/(cellValue);
                        }
                        else
                        {
                            dataGridViewRow.Cells[col - 1].Value = worksheet.Cells[row, col].Text;
                        }
                    }

                    // Check for matching store names and add them to the "ExtractedStoreName" column
                    string rfpDescription = dataGridViewRow.Cells[0].Value?.ToString();
                    string extractedStoreNames = GetMatchingStoreNames(rfpDescription);
                    dataGridViewRow.Cells[dataGridView1.Columns["ExtractedStoreName"].Index].Value = extractedStoreNames;

                    // Check for matching TOCs and add them to the "ExtractedTOCs" column
                    string extractedTOCs = GetMatchingTOCs(rfpDescription);
                    dataGridViewRow.Cells[dataGridView1.Columns["ExtractedTOCs"].Index].Value = extractedTOCs;

                    dataGridView1.Rows.Add(dataGridViewRow);
                }
            }
        }

        // Function to find all matching store names in dataGridView2 based on the RFP Description
        private string GetMatchingStoreNames(string rfpDescription)
        {
            if (string.IsNullOrEmpty(rfpDescription)) return string.Empty;

            string matchedStores = string.Empty;

            foreach (DataGridViewRow storeRow in dataGridView2.Rows)
            {
                if (storeRow.IsNewRow) continue;

                string storeName = storeRow.Cells["StoreName"].Value?.ToString();
                string storeCode = storeRow.Cells["StoreCode"].Value?.ToString();

                if (!string.IsNullOrEmpty(storeName) && rfpDescription.Contains(storeName))
                {
                    // Concatenate StoreCode and StoreName with a dash in between
                    matchedStores += storeCode + " - " + storeName + ", ";
                }
            }

            // Remove the trailing comma and space, if any
            return matchedStores.TrimEnd(',', ' ');
        }



        private string GetMatchingTOCs(string rfpDescription)
        {
            if (string.IsNullOrEmpty(rfpDescription)) return string.Empty;

            HashSet<string> matchedTOCs = new HashSet<string>();
            List<string> tempMatches = new List<string>(); // Temporary list to hold all matches

            foreach (DataGridViewRow tocRow in dataGridView3.Rows)
            {
                if (tocRow.IsNewRow) continue;

                string tocName = tocRow.Cells[0].Value?.ToString();
                if (!string.IsNullOrEmpty(tocName))
                {
                    
                    string pattern;
                    if (tocName.Contains("(") || tocName.Contains(")"))
                    {
                        pattern = Regex.Escape(tocName);
  
                    }
                    else
                    {
                        pattern = @"\b" + Regex.Escape(tocName) + @"\b(?:\s*\(.*\))?";
                    }

                    if (Regex.IsMatch(rfpDescription, pattern, RegexOptions.IgnoreCase))
                    {
                        tempMatches.Add(tocName); // Add matched TOC to temporary list
                    }
                }
            }

            // Now filter tempMatches to only include the most specific versions
            foreach (string match in tempMatches)
            {
                // Check if the match contains parentheses
                if (match.Contains("("))
                {
                    // Add the more specific match if it's not already in matchedTOCs
                    if (!matchedTOCs.Contains(match))
                    {
                        matchedTOCs.Add(match);
                        Console.WriteLine($"TOCNAME: {match}");
                    }
                }
                else
                {
                    // If the less specific match exists, do not add it if a more specific match is already present
                    if (!matchedTOCs.Any(m => m.StartsWith(match + " ") || m.Equals(match + " (")))
                    {
                        matchedTOCs.Add(match);
                    }
                }
            }

            return string.Join(", ", matchedTOCs);
        }




        private string ExtractRfpNumber(string input)
        {
            int startIndex = input.IndexOf("RFP#LAWHO");
            if (startIndex == -1) return input; // Return original if "RFP#" not found

            int endIndex = startIndex + 4; // Move past "RFP#"

            // Continue until we hit a space or non-word character
            while (endIndex < input.Length && (char.IsLetterOrDigit(input[endIndex]) || input[endIndex] == '-' || char.IsWhiteSpace(input[endIndex])))
            {
                endIndex++;
            }

            return input.Substring(startIndex, endIndex - startIndex).Trim();
        }

        private void LoadComboBoxItems()
        {
            // Add items to the ComboBox
            comboBox1.Items.Add("CDJ");
            comboBox1.Items.Add("APV");
            comboBox1.Items.Add("Inventory_Transaction");
            //comboBox1.Items.Add("ChargeTo");
        }

        private void UploadDb_Click(object sender, EventArgs e)
        {
            string selectedType = comboBox1.SelectedItem?.ToString();
            if (selectedType == "APV")
            {
                UploadToAPV();
            }
            else if (selectedType == "CDJ")
            {
                UploadToCDJ();
            }
            else if (selectedType == "Inventory_Transaction")
            {
                UploadToInventory_Transactions();
            }
            //else if (selectedType == "ChargeTo")
            //{
            //    UploadToTOCs();
            //}
            else
            {
                MessageBox.Show("Please select a valid option.");
            }
        }

        private void UploadToAPV()
        {
            int insertedCount = 0; // Keep track of successful insertions

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                string invoiceDescription =(row.Cells[0].Value?.ToString());
                string voucher = row.Cells[1].Value?.ToString();
                string date = row.Cells[2].Value?.ToString();
                string invoiceAccount = row.Cells[3].Value?.ToString();
                string duedate = row.Cells[4].Value?.ToString();

                InsertIntoAPV(invoiceDescription, voucher, date, invoiceAccount, duedate);
                insertedCount++; // Increment the count for every successful insertion
            }

            // Show a pop-up message once after all rows are processed
            if (insertedCount > 0)
            {
                MessageBox.Show($"{insertedCount} APV records inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void UploadToCDJ()
        {
            int insertedCount = 0; // Keep track of successful insertions

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                string rfp =(row.Cells[0].Value?.ToString());
                string voucher = row.Cells[2].Value?.ToString();
                string lastSettleVoucher = row.Cells[1].Value?.ToString();
                string approvedDate = row.Cells[3].Value?.ToString();
                string closed = row.Cells[4].Value?.ToString();
                string storeName = row.Cells[5].Value?.ToString();
                string tocname = row.Cells[6].Value?.ToString();

                InsertIntoCDJ(rfp, voucher, lastSettleVoucher, approvedDate, closed, storeName, tocname);
                insertedCount++; // Increment the count for every successful insertion
            }

            // Show a pop-up message once after all rows are processed
            if (insertedCount > 0)
            {
                MessageBox.Show($"{insertedCount} CDJ records inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void UploadToInventory_Transactions()
        {
            int insertedCount = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                if (row.Cells.Count < 21) continue; // Adjust based on the expected number of columns

                string Item_number = row.Cells[0].Value?.ToString();
                string Physical_date = row.Cells[1].Value?.ToString();
                string Financial_date = row.Cells[2].Value?.ToString();
                string Reference = row.Cells[3].Value?.ToString();
                string ItemNumber = row.Cells[4].Value?.ToString();
                string PhysicalDate = row.Cells[5].Value?.ToString();
                string FinancialDate = row.Cells[6].Value?.ToString();
                string Referenceaa = row.Cells[7].Value?.ToString();
                string Number = row.Cells[8].Value?.ToString();
                string Receipt = row.Cells[9].Value?.ToString();
                string Issue = row.Cells[10].Value?.ToString();
                string Unit = row.Cells[11].Value?.ToString();
                string Site = row.Cells[12].Value?.ToString();
                string Warehouse = row.Cells[13].Value?.ToString();
                string Quantity = row.Cells[14].Value?.ToString();
                string Location = row.Cells[15].Value?.ToString();
                string CostAmount = row.Cells[16].Value?.ToString();
                string Adjustment = row.Cells[17].Value?.ToString();
                string FinancialCostAmount = row.Cells[18].Value?.ToString();
                string PhysicalCostAmount = row.Cells[19].Value?.ToString();
                string ModifiedDateTime = row.Cells[20].Value?.ToString();

                InsertIntoInventory_Transactions(
                    Item_number,
                    Physical_date,
                    Financial_date,
                    Reference,
                    ItemNumber,
                    PhysicalDate,
                    FinancialDate,
                    Referenceaa,
                    Number,
                    Receipt,
                    Issue,
                    Unit,
                    Site,
                    Warehouse,
                    Quantity,
                    Location,
                    CostAmount,
                    Adjustment,
                    FinancialCostAmount,
                    PhysicalCostAmount,
                    ModifiedDateTime
                );
                insertedCount++;
            }

            if (insertedCount > 0)
            {
                MessageBox.Show($"{insertedCount} Inventory_Transactions records inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }



        private void InsertIntoAPV(string invoiceDescription, string voucher, string date, string invoiceAccount, string duedate)
        {
            string extractedRFPDescription = ExtractNumberFromRFP(invoiceDescription);

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionSQL"].ConnectionString))
            {
                string query = "INSERT INTO APV_Extraction (Invoice_description, Voucher, Date, Invoice_account, DueDate) VALUES (@InvoiceDescription, @Voucher, @Date, @InvoiceAccount, @DueDate)";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
          

                    command.Parameters.AddWithValue("@InvoiceDescription", extractedRFPDescription);
                    command.Parameters.AddWithValue("@Voucher", voucher);
                    command.Parameters.AddWithValue("@Date", date);
                    command.Parameters.AddWithValue("@InvoiceAccount", invoiceAccount);
                    command.Parameters.AddWithValue("@DueDate", duedate);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
        private void InsertIntoCDJ(string rfp, string voucher, string lastSettleVoucher, string approvedDate, string closed, string storeName, string tocname)
        {
            string extractedRFP = ExtractNumberFromRFP(rfp);

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionSQL"].ConnectionString))
            {
                string query = "INSERT INTO CDJ_Extraction (RFP#, Voucher, LastSettleVoucher, ApprovedDate, Closed, StoreName, TOCNAME) VALUES (@RFP, @Voucher, @LastSettleVoucher, @ApprovedDate, @Closed, @StoreName, @TOCNAME)";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@RFP", extractedRFP);
                    command.Parameters.AddWithValue("@Voucher", voucher);
                    command.Parameters.AddWithValue("@LastSettleVoucher", lastSettleVoucher);
                    command.Parameters.AddWithValue("@ApprovedDate", approvedDate);
                    command.Parameters.AddWithValue("@Closed", closed);
                    command.Parameters.AddWithValue("@StoreName", storeName);
                    command.Parameters.AddWithValue("@TOCNAME", tocname);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

        private void InsertIntoInventory_Transactions(
      string Item_number,
      string Physical_date,
      string Financial_date,
      string Reference,
      string ItemNumber,
      string PhysicalDate,
      string FinancialDate,
      string Referenceaa,
      string Number,
      string Receipt,
      string Issue,
      string Unit,
      string Site,
      string Warehouse,
      string Quantity,
      string Location,
      string CostAmount,
      string Adjustment,
      string FinancialCostAmount,
      string PhysicalCostAmount,
      string ModifiedDateTime)
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionSQL2"].ConnectionString))
            {
              
                string query = "INSERT INTO Inventory_Transaction (Item_number, Physical_date, Financial_date, Reference, ItemNumber, PhysicalDate, FinancialDate, Referenceaa, Number, Receipt, Issue, Unit, Site, Warehouse, Quantity, Location, CostAmount, Adjustment, FinancialCostAmount, PhysicalCostAmount, ModifiedDateTime) VALUES (@Item_number, @Physical_date, @Financial_date, @Reference, @ItemNumber, @PhysicalDate, @FinancialDate, @Referenceaa, @Number, @Receipt, @Issue, @Unit, @Site, @Warehouse, @Quantity, @Location, @CostAmount, @Adjustment, @FinancialCostAmount, @PhysicalCostAmount, @ModifiedDateTime)";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Item_number", Item_number);
                    command.Parameters.AddWithValue("@Physical_date", Physical_date);
                    command.Parameters.AddWithValue("@Financial_date", Financial_date);
                    command.Parameters.AddWithValue("@Reference", Reference);
                    command.Parameters.AddWithValue("@ItemNumber", ItemNumber);
                    command.Parameters.AddWithValue("@PhysicalDate", PhysicalDate);
                    command.Parameters.AddWithValue("@FinancialDate", FinancialDate);
                    command.Parameters.AddWithValue("@Referenceaa", Referenceaa);
                    command.Parameters.AddWithValue("@Number", Number);
                    command.Parameters.AddWithValue("@Receipt", Receipt);
                    command.Parameters.AddWithValue("@Issue", Issue);
                    command.Parameters.AddWithValue("@Unit", Unit);
                    command.Parameters.AddWithValue("@Site", Site);
                    command.Parameters.AddWithValue("@Warehouse", Warehouse);
                    command.Parameters.AddWithValue("@Quantity", Quantity);
                    command.Parameters.AddWithValue("@Location", Location);
                    command.Parameters.AddWithValue("@CostAmount", CostAmount);
                    command.Parameters.AddWithValue("@Adjustment", Adjustment);
                    command.Parameters.AddWithValue("@FinancialCostAmount", FinancialCostAmount);
                    command.Parameters.AddWithValue("@PhysicalCostAmount", PhysicalCostAmount);
                    command.Parameters.AddWithValue("@ModifiedDateTime", ModifiedDateTime);
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }


        private string ExtractNumberFromRFP(string rfpValue)
        {
            int startIndex = rfpValue.IndexOf("LAWHO-");
            if (startIndex != -1)
            {
                startIndex += "LAWHO-".Length; // Move index after "LAWHO-"

                // Calculate the maximum possible length after "LAWHO-"
                int remainingLength = rfpValue.Length - startIndex;
                int length = Math.Min(8, remainingLength); // Ensure length doesn't exceed available characters

                if (length > 0) // Check if there's enough length to extract
                {
                    string extractedNumber = rfpValue.Substring(startIndex, length);

                    // Remove leading '1' and all following '0's until a non-zero digit is found
                    int i = 0;

                    // Check if the first character is '1'
                    if (extractedNumber[i] == '1')
                    {
                        i++; // Move past the first '1'
                        while (i < extractedNumber.Length && extractedNumber[i] == '0')
                        {
                            i++; // Skip all '0's
                        }
                    }

                    // Return the remaining part, including '0's or more
                    return extractedNumber.Substring(i);
                }
            }
            return string.Empty; // Return empty if "LAWHO-" not found or insufficient length
        }




        private void LoadStoreNamesIntoGridView()
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionSQL"].ConnectionString))
            {
                connection.Open();
                string query = "SELECT StoreCode, StoreName FROM Stores"; // Adjust query as necessary
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();

                dataGridView2.Columns.Clear();
                dataGridView2.Rows.Clear();
                dataGridView2.Columns.Add("StoreName", "Store Name");
                dataGridView2.Columns.Add("StoreCode", "Store Code");

                while (reader.Read())
                {
                    dataGridView2.Rows.Add(reader["StoreName"].ToString(), reader["StoreCode"].ToString());
                }
            }
        }

        private void LoadTOCNamesIntoGridView()
        {
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionSQL"].ConnectionString))
            {
                connection.Open();
                string query = "SELECT DISTINCT TOCName FROM TOCs"; // Adjust query as necessary
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();

                dataGridView3.Columns.Clear();
                dataGridView3.Rows.Clear();
                dataGridView3.Columns.Add("TOCName", "TOC Name");

                while (reader.Read())
                {
                    dataGridView3.Rows.Add(reader["TOCName"].ToString());
                }
            }
        }
    }
}
