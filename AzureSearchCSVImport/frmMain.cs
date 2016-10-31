using CsvHelper;
using Microsoft.Azure.Search;
using Microsoft.Azure.Search.Models;
using Microsoft.Spatial;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AzureSearchCSVImport
{
    public partial class frmMain : Form
    {
        private static SearchServiceClient serviceClient;

        public frmMain()
        {
            InitializeComponent();
            ConfigureSearchDataGrid();
            SetStatus("Ready.  Please choose a CSV file to import.");
            btnImportData.Enabled = false;
            btnCreateIndex.Enabled = false;
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            try
            {
                List<ColumnMetadata> ColumnMetaDataList = new List<ColumnMetadata>();
                openFileDialog1.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt";
                DialogResult result = openFileDialog1.ShowDialog();
                int rowCount = 0;
                if (result == DialogResult.OK)
                {
                    btnImportData.Enabled = false;
                    btnCreateIndex.Enabled = false;
                    SetStatus("Parsing CSV file...");
                    
                    txtFileName.Text = openFileDialog1.FileName;

                    using (var sr = new StreamReader(txtFileName.Text))
                    {
                        var csv = new CsvReader(sr);
                        int colCount = 0;
                        bool rowLimit = false;
                        searchDataGrid.Rows.Clear();
                        searchDataGrid.Refresh();
                        csvDataGrid.Rows.Clear();
                        csvDataGrid.Refresh();
                        csvDataGrid.ColumnCount = 0;;

                        // Get the columnn names and how many columns of data are there
                        while ((csv.Read()) && (rowLimit == false))
                        {
                            if (rowCount == 0)
                            {

                                bool moreColumns = true;
                                while (moreColumns)
                                {
                                    try
                                    {
                                        csvDataGrid.ColumnCount += 1;
                                        string stringField = csv.GetField<string>(colCount);
                                        csvDataGrid.Columns[csvDataGrid.ColumnCount - 1].Name = stringField;
                                        ColumnMetaDataList.Add(new ColumnMetadata { ColumnName = stringField });
                                        colCount++;
                                    }
                                    catch (Exception)
                                    {
                                        // This will be called when there is no more columns
                                        moreColumns = false;
                                    }
                                }
                            }
                            else
                            {
                                List<string> rowData = new List<string>();
                                for (int col = 0; col < ColumnMetaDataList.Count; col++)
                                {
                                    string stringField = csv.GetField<string>(col);
                                    rowData.Add(stringField);
                                    DataType dt = EvaluateDataType(stringField);
                                    // If it is a string I can stop here:
                                    if (dt == DataType.String)
                                    {
                                        ColumnMetaDataList[col].ColumnDataType = dt;
                                        if (stringField != null)
                                            if (ColumnMetaDataList[col].ColumnMaxLen < stringField.Length)
                                                ColumnMetaDataList[col].ColumnMaxLen = stringField.Length;
                                    }
                                    else if (dt == DataType.Int32)
                                    {
                                        if ((ColumnMetaDataList[col].ColumnDataType != DataType.String) &&
                                            (ColumnMetaDataList[col].ColumnDataType != DataType.Int64) &&
                                            (ColumnMetaDataList[col].ColumnDataType != DataType.Double) &&
                                            (ColumnMetaDataList[col].ColumnDataType != DataType.DateTimeOffset))
                                            ColumnMetaDataList[col].ColumnDataType = dt;
                                    }
                                    else if (dt == DataType.Int64)
                                    {
                                        if ((ColumnMetaDataList[col].ColumnDataType != DataType.String) &&
                                            (ColumnMetaDataList[col].ColumnDataType != DataType.Double) &&
                                            (ColumnMetaDataList[col].ColumnDataType != DataType.DateTimeOffset))
                                            ColumnMetaDataList[col].ColumnDataType = dt;
                                    }
                                    else if (dt == DataType.Double)
                                    {
                                        if ((ColumnMetaDataList[col].ColumnDataType != DataType.String) &&
                                            ColumnMetaDataList[col].ColumnDataType != DataType.DateTimeOffset)
                                            ColumnMetaDataList[col].ColumnDataType = dt;
                                    }
                                    else if (ColumnMetaDataList[col].ColumnDataType != DataType.String)
                                    {
                                        ColumnMetaDataList[col].ColumnDataType = dt;
                                    }
                                }
                                csvDataGrid.Rows.Add(rowData.ToArray());

                            }


                            rowCount++;
                            if (rowCount == 100)
                                rowLimit = true;

                        }
                    }

                    rowCount = 0;
                    foreach (ColumnMetadata cm in ColumnMetaDataList)
                    {
                        FieldProperties fp = GetDefaultFieldProperties(cm.ColumnDataType, cm.ColumnMaxLen);
                        string isKey = "False";
                        if (rowCount == 0)
                            isKey = "True";

                        string dt = "Edm.String";
                        if (cm.ColumnDataType == DataType.Int32)
                            dt = "Edm.Int32";
                        else if (cm.ColumnDataType == DataType.Int64)
                            dt = "Edm.Int64";
                        else if (cm.ColumnDataType == DataType.Double)
                            dt = "Edm.Double";
                        else if (cm.ColumnDataType == DataType.Boolean)
                            dt = "Edm.Boolean";
                        else if (cm.ColumnDataType == DataType.DateTimeOffset)
                            dt = "Edm.DateTimeOffset";

                        string[] row = new string[] { ConvertToCamelCase(cm.ColumnName), dt, isKey, fp.Retreiveable, fp.Filterable, fp.Sortable, fp.Facetable, fp.Searchable };
                        searchDataGrid.Rows.Add(row);
                        rowCount++;

                    }
                }
            }
            catch (Exception ex)
            {
                SetStatus("Error: " + ex.Message);
                return;
            }
            SetStatus("Ready.");
            btnImportData.Enabled = true;
            btnCreateIndex.Enabled = true;

        }

        private void ConfigureSearchDataGrid()
        {
            // Configure the search data grid
            searchDataGrid.ColumnCount = 1;
            searchDataGrid.Columns[0].Name = "Field Name";

            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Data Type";
            cmb.Name = "cmbDataType";
            cmb.MaxDropDownItems = 7;
            cmb.Items.Add("Edm.String");
            cmb.Items.Add("Edm.Boolean");
            cmb.Items.Add("Edm.Int32");
            cmb.Items.Add("Edm.Int64");
            cmb.Items.Add("Edm.Double");
            cmb.Items.Add("Edm.DateTimeOffset");
            cmb.Items.Add("Edm.GeographyPoint");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Key";
            cmb.Name = "cmbKey";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Retrieveable";
            cmb.Name = "cmbRetrieveable";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Filterable";
            cmb.Name = "cmbFilterable";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Sortable";
            cmb.Name = "cmbSortable";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Facetable";
            cmb.Name = "cmbFacetable";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

            cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Searchable";
            cmb.Name = "cmbSearchable";
            cmb.MaxDropDownItems = 2;
            cmb.Items.Add("True");
            cmb.Items.Add("False");
            searchDataGrid.Columns.Add(cmb);

        }

        private string ConvertToCamelCase(string input)
        {
            string[] split = input.Split(' ');
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < split.Count(); i++)
            {
                if (i == 0)
                    sb.Append(split[i].ToLower());
                else
                    sb.Append(split[i]);
            }

            return sb.ToString();
        }

        public FieldProperties GetDefaultFieldProperties(DataType FieldType, Int32 MaxFieldLength = 0)
        {
            // Figure out the best default properties
            // Max field len helps to tell if it should be filterable, facetable, etc
            FieldProperties fp = new FieldProperties();
            fp.FieldType = FieldType;
            fp.Retreiveable = "True";

            if (FieldType ==  DataType.String)
            {
                fp.Searchable = "True";
                if (MaxFieldLength <= 30)
                {
                    fp.Filterable = "True";
                    fp.Sortable = "True";
                    fp.Facetable = "True";
                }
                else
                {
                    fp.Filterable = "False";
                    fp.Sortable = "False";
                    fp.Facetable = "False";
                }
            }
            else if (FieldType == DataType.Boolean)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "True";
                fp.Facetable = "True";
            }
            else if (FieldType == DataType.Int32)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "True";
                fp.Facetable = "True";
            }
            else if (FieldType == DataType.Int64)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "True";
                fp.Facetable = "True";
            }
            else if (FieldType == DataType.Double)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "True";
                fp.Facetable = "True";
            }
            else if (FieldType == DataType.DateTimeOffset)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "True";
                fp.Facetable = "True";
            }
            else if (FieldType == DataType.GeographyPoint)
            {
                fp.Searchable = "False";
                fp.Filterable = "True";
                fp.Sortable = "False";
                fp.Facetable = "True";
            }

            return fp;
        }

        private void searchDataGrid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // Handle the case where they change the field type
            if (searchDataGrid.CurrentCell.ColumnIndex == 1 && e.Control is ComboBox)
            {
                ComboBox comboBox = e.Control as ComboBox;
                comboBox.SelectedIndexChanged += FieldTypeSelectionChanged;
            }
        }


        private void FieldTypeSelectionChanged(object sender, EventArgs e)
        {
            // Take the newly selected field type and update the properties
            var currentcell = searchDataGrid.CurrentCellAddress;
            var sendingCB = sender as DataGridViewComboBoxEditingControl;
            string FieldType = sendingCB.EditingControlFormattedValue.ToString();
            DataType dt = null;
            if (FieldType == "Edm.String")
                dt = DataType.String;
            else if (FieldType == "Edm.Int32")
                dt = DataType.Int32;
            else if (FieldType == "Edm.Int64")
                dt = DataType.Int64;
            else if (FieldType == "Edm.Double")
                dt = DataType.Double;
            else if (FieldType == "Edm.DateTimeOffset")
                dt = DataType.DateTimeOffset;

            FieldProperties fp = GetDefaultFieldProperties(dt);

            DataGridViewComboBoxCell cel = (DataGridViewComboBoxCell)searchDataGrid.Rows[currentcell.Y].Cells[3];
            cel.Value = fp.Retreiveable;
            cel = (DataGridViewComboBoxCell)searchDataGrid.Rows[currentcell.Y].Cells[4];
            cel.Value = fp.Filterable;
            cel = (DataGridViewComboBoxCell)searchDataGrid.Rows[currentcell.Y].Cells[5];
            cel.Value = fp.Sortable;
            cel = (DataGridViewComboBoxCell)searchDataGrid.Rows[currentcell.Y].Cells[6];
            cel.Value = fp.Facetable;
            cel = (DataGridViewComboBoxCell)searchDataGrid.Rows[currentcell.Y].Cells[7];
            cel.Value = fp.Searchable;

        }

        private void btnCreateIndex_Click(object sender, EventArgs e)
        {
            // Use the defined schema to create a new index
            if (ValidateInputs() == false)
                return;
            // Check if the Index already exists
            SetStatus("Deleting index...");
            if (DeleteIndex() == false)
                return;
            CreateIndex();

        }

        private void SetStatus(string status)
        {
            toolStripStatusLabel1.Text = status;
            statusStrip1.Refresh();
        }

        private bool DeleteIndex()
        {
            // Delete the index if it exists
            try
            {
                serviceClient = new SearchServiceClient(txtSearchService.Text, new SearchCredentials(txtApiKey.Text));
                if (serviceClient.Indexes.Exists(txtIndexName.Text))
                    serviceClient.Indexes.Delete(txtIndexName.Text);
            }
            catch (Exception ex)
            {
                SetStatus("Error deleting index: " + ex.Message);
                return true;
            }
            SetStatus("Index deleted, re-creating...");
            return true;
        }

        private void CreateIndex()
        {

            try
            {
                serviceClient = new SearchServiceClient(txtSearchService.Text, new SearchCredentials(txtApiKey.Text));

                // Set a keyfield
                List<Field> FieldList = new List<Field>();
                Field f = new Field("keyField", DataType.String) { IsKey = true, IsFacetable = false, IsFilterable = false, IsSearchable = false, IsRetrievable = true, IsSortable = false };
                FieldList.Add(f);

                for (int i = 0; i < searchDataGrid.RowCount - 1; i++)
                {
                    string FieldName = searchDataGrid[0, i].Value.ToString();
                    DataType dt = null;
                    string FieldType = searchDataGrid[1, i].Value.ToString();
                    if (FieldType == "Edm.String")
                        dt = DataType.String;
                    else if (FieldType == "Edm.Boolean")
                        dt = DataType.Boolean;
                    else if (FieldType == "Edm.Int32")
                        dt = DataType.Int32;
                    else if (FieldType == "Edm.Int64")
                        dt = DataType.Int64;
                    else if (FieldType == "Edm.Double")
                        dt = DataType.Double;
                    else if (FieldType == "Edm.DateTimeOffset")
                        dt = DataType.DateTimeOffset;
                    else if (FieldType == "Edm.GeographyPoint")
                        dt = DataType.GeographyPoint;

                    f = new Field(FieldName, dt);

                    f.IsKey = false;
                    //if (i == 0)
                    //    f.IsKey = true;

                    f.IsRetrievable = false;
                    if (searchDataGrid[3, i].Value.ToString() == "True")
                        f.IsRetrievable = true;

                    f.IsFilterable = false;
                    if (searchDataGrid[4, i].Value.ToString() == "True")
                        f.IsFilterable = true;

                    f.IsSortable = false;
                    if (searchDataGrid[5, i].Value.ToString() == "True")
                        f.IsSortable = true;

                    f.IsFacetable = false;
                    if (searchDataGrid[6, i].Value.ToString() == "True")
                        f.IsFacetable = true;

                    f.IsSearchable = false;
                    if (searchDataGrid[7, i].Value.ToString() == "True")
                        f.IsSearchable = true;

                    FieldList.Add(f);

                }

                var definition = new Index()
                {
                    Name = txtIndexName.Text,
                    Fields = FieldList.ToArray()
                };

                serviceClient.Indexes.Create(definition);
                SetStatus("Index created successfully");

            }
            catch (Exception ex)
            {
                SetStatus("Error: " + ex.Message.ToString());
            }

        }


        private void btnImportData_Click(object sender, EventArgs e)
        {
            // Upload all the data from specified CSV
            if (ValidateInputs() == false)
                return;

            serviceClient = new SearchServiceClient(txtSearchService.Text, new SearchCredentials(txtApiKey.Text));
            ISearchIndexClient indexClient = serviceClient.Indexes.GetClient(txtIndexName.Text);

            //Document doc1 = new Document();
            //doc1.Add("vm", "A0 VM Windows");
            //List<IndexAction> actions1 = new List<IndexAction>();
            //actions1.Add(IndexAction.MergeOrUpload(doc1));
            //indexClient.Documents.Index(new IndexBatch(actions1));


            // Take the data from the CSV and import into Azure Search in batches of 1000
            using (var sr = new StreamReader(txtFileName.Text))
            {
                var csv = new CsvReader(sr);
                int colCount = 0;
                int rowCount = 0;
                int batchCounter = 0;
                bool errorFound = false;
                try
                {
                    SetStatus("Uploading documents...");
                    List<IndexAction> actions = new List<IndexAction>();
                    while ((csv.Read()) && (errorFound == false))
                    {
                        // Set cursor as hourglass
                        Cursor.Current = Cursors.WaitCursor;

                        rowCount += 1;

                        Document doc = new Document();
                        // Skip first row which is header
                        if (rowCount != 1)
                        {
                            colCount = 0;
                            bool moreData = true;
                            while (moreData == true)
                            {
                                try
                                {
                                    string dataType = searchDataGrid[1, colCount].Value.ToString();
                                    string dataValue = csv.GetField<String>(colCount);
                                    string keyValue = searchDataGrid[0, colCount].Value.ToString();
                                    string isKey = searchDataGrid[2, colCount].Value.ToString();
                                    if (isKey == "True")
                                    {
                                        // Ensure that the key is not empty which is a bad key field
                                        if (dataValue.Length == 0)
                                        {
                                            MessageBox.Show("A blank value was found in the key field for row number " + rowCount + ", please use a different field for the key.", "Error");
                                            errorFound = true;
                                            break;
                                        }
                                        doc.Add("keyField", Convert.ToBase64String(Encoding.UTF8.GetBytes(dataValue)));
                                    }
                                    if (dataType == "Edm.String")
                                        doc.Add(keyValue, dataValue);
                                    else if (dataType == "Edm.Boolean")
                                        doc.Add(keyValue, Convert.ToBoolean(dataValue.Replace(",", "").Replace(" ", "")));
                                    else if (dataType == "Edm.Int32")
                                        doc.Add(keyValue, Convert.ToInt32(dataValue.Replace(",", "").Replace(" ", "")));
                                    else if (dataType == "Edm.Int64")
                                        doc.Add(keyValue, Convert.ToInt64(dataValue.Replace(",", "").Replace(" ", "")));
                                    else if (dataType == "Edm.Double")
                                        doc.Add(keyValue, Convert.ToDouble(dataValue.Replace(",", "").Replace(" ", "")));
                                    else if (dataType == "Edm.DateTimeOffset")
                                        doc.Add(keyValue, Convert.ToDateTime(dataValue.Replace(",", "").Replace(" ", "")));
                                    //else if (dataType == "Edm.GeographyPoint")
                                    //    doc.Add(keyValue, csv.GetField<Boolean>(colCount));

                                    colCount++;
                                }
                                catch (Exception)
                                {
                                    // This will be called when there is no more columns to process
                                    moreData = false;
                                }
                            }

                            actions.Add(IndexAction.MergeOrUpload(doc));
                            batchCounter++;
                            Application.DoEvents();     // Allow the user to do things like minimize or cancel
                            if ((batchCounter == Convert.ToInt32(txtBatchSize.Text)) && (errorFound == false))
                            {
                                indexClient.Documents.Index(new IndexBatch(actions));
                                actions = new List<IndexAction>();
                                batchCounter = 0;
                                SetStatus((rowCount - 1).ToString() + " documents uploaded...");
                            }

                        }
                    }
                    if ((batchCounter > 0) && (errorFound == false))
                        indexClient.Documents.Index(new IndexBatch(actions));

                }
                catch (Exception ex)
                {
                    SetStatus("Error: " + ex.Message);
                    return;
                }
                SetStatus("Upload Complete. " + (rowCount - 1).ToString() + " documents uploaded.");
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }
        }

        private DataType EvaluateDataType(string value)
        {
            // Determine what the most likely datatype is

            // If I remove the , and spaces, is it an double?
            try
            {
                if (value != null)
                {
                    string valueToTest = value.Replace(",", "").Replace(" ", "");
                    double n;
                    // From excel it is common to have numbers that look like "2.14E+12" which I will just make a string
                    if ((double.TryParse(valueToTest, out n)) && (valueToTest.IndexOf(".") > -1) && (valueToTest.IndexOf("E+") == -1))
                        return DataType.Double;

                    Int64 i;
                    if (Int64.TryParse(valueToTest, out i))
                    {
                        if ((i > 214748364) || (i < -214748364))
                            return DataType.Int64;
                        else
                            return DataType.Int32;
                    }

                    DateTimeOffset d;
                    if (DateTimeOffset.TryParse(valueToTest, out d))
                        return DataType.DateTimeOffset;
                }
            }
            catch (Exception)
            {
                return DataType.String;
            }
            return DataType.String;


        }

        private void txtBatchSize_TextChanged(object sender, EventArgs e)
        {
            

        }

        private void txtBatchSize_Leave(object sender, EventArgs e)
        {
            // Verify it is between 0 and 1000
            try
            {
                Int32 i;
                if (Int32.TryParse(txtBatchSize.Text, out i))
                {
                    if ((Convert.ToInt32(txtBatchSize.Text) < 1) || (Convert.ToInt32(txtBatchSize.Text) > 1000))
                    {
                        MessageBox.Show("Please choose a batch size between 1 and 1000.");
                        txtBatchSize.Text = "1000";
                    }
                } else
                {
                    MessageBox.Show("Please choose a batch size between 1 and 1000.");
                    txtBatchSize.Text = "1000";
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please choose a batch size between 1 and 1000.");
                txtBatchSize.Text = "1000";
            }
        }

        private void txtSearchService_Leave(object sender, EventArgs e)
        {
            if (txtSearchService.Text.ToLower().IndexOf(".search.windows.net") > -1)
            {
                MessageBox.Show("Full URL is not required (.search.windows.net) in the service name.", "Info");
                txtSearchService.Text = txtSearchService.Text.ToLower().Replace(".search.windows.net", "");
                return;
            }
        }

        private bool ValidateInputs()
        {
            string errorMsg = string.Empty;
            if (txtSearchService.Text.Length == 0)
                errorMsg = "Please enter a valid search service name.";
            else if (txtApiKey.Text.Length == 0)
                errorMsg = "Please enter a valid API key.  This should be an admin key and is accessible from the Azure portal for your Azure Search service.";
            else if (txtIndexName.Text.Length == 0)
                errorMsg = "Please enter a valid index name.";

            if (errorMsg != string.Empty)
            {
                MessageBox.Show(errorMsg, "Error");
                SetStatus(errorMsg);
                return false;
            }
            return true;
        }
    }

    public class ColumnMetadata
    {
        public string ColumnName { get; set; }
        public DataType ColumnDataType { get; set; }
        public Int32 ColumnMaxLen { get; set; }
    }
}