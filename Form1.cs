using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Xml;

namespace BNPPartsCatalog
{
    public partial class frmBNPPartsCatalog : Form
    {
        private string userGuide = @"..\BNPUserGuide.html";
        private string filePath = @"..\Inventory.xml";
        private DataSet dsInventory = new DataSet();
        private DataTable dtProducts;
        private DataRow drRecord;

        //- This will be the name of our Table in the XML file:
        private string sTableName = "Product";

        //- These are the names of the columns that make up a row in our XML file:
        private string sMakeColumn = "Make";
        private string sModelColumn = "Model";
        private string sPartNumberColumn = "PartNumber";
        private string sPriceColumn = "Price";
        private string sManufacturerColumn = "Manufacturer";
        private string sManufacturerIDColumn = "ManufacturerID";
        private string sItemNameColumn = "ItemName";
        private string sItemTypeColumn = "ItemType";
        private string sConditionColumn = "Condition";
        private string sSerialNumberColumn = "SerialNumber";
        private string sLocation = "Location";
        private string sExpirationColumn = "ExpirationDate";
        private string sItemDescriptionColumn = "Description";

        private BindingSource bsInfo = new BindingSource();

        public frmBNPPartsCatalog()
        {
            InitializeComponent();

            //- Set the name of our DataTable to the table name in our XML file:
            dtProducts = new DataTable(sTableName);

            //- Add the columns that we declared above to our DataTable:
            dtProducts.Columns.Add(sMakeColumn, typeof(string));
            dtProducts.Columns.Add(sModelColumn, typeof(string));
            dtProducts.Columns.Add(sPartNumberColumn, typeof(string));
            dtProducts.Columns.Add(sPriceColumn, typeof(string));
            dtProducts.Columns.Add(sManufacturerColumn, typeof(string));
            dtProducts.Columns.Add(sManufacturerIDColumn, typeof(string));
            dtProducts.Columns.Add(sItemNameColumn, typeof(string));
            dtProducts.Columns.Add(sItemTypeColumn, typeof(string));
            dtProducts.Columns.Add(sConditionColumn, typeof(string));
            dtProducts.Columns.Add(sSerialNumberColumn, typeof(string));
            dtProducts.Columns.Add(sLocation, typeof(string));
            dtProducts.Columns.Add(sExpirationColumn, typeof(DateTime));
            dtProducts.Columns.Add(sItemDescriptionColumn, typeof(string));

            //- Add the DataTable to the DataSet:
            dsInventory.Tables.Add(dtProducts);

            getDataSource();

            populateFields();
            populateGrid();
            bindNagivation();

            checkForExpiredItems();
        }

        private void getDataSource()
        {
            //- Display the number of records in the XML file:
            XmlDocument doc = new XmlDocument();
            doc.Load(filePath);
            XmlNodeList elemList = doc.GetElementsByTagName(sTableName);
            //MessageBox.Show("The total number of records in the database is: " + elemList.Count);

            dsInventory.Clear();

            //- If the xml file exists:
            if (File.Exists(filePath))
            {
                //- Then read the xml file into a DataSet:
                dsInventory.ReadXml(filePath);
                dsInventory.AcceptChanges();
            }

            /* Set the BindingSource to the DataSource, then in order to get the Binding Navigator to work correctly, 
             * Set the DataMember of the BindingSource to the Table Name.*/
            bsInfo.DataSource = dsInventory;
            bsInfo.DataMember = sTableName;
        }

        private void populateFields()
        {
            //- Bind the Form Fields to the Binding Source:
            this.tbMake.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sMakeColumn, true));
            this.tbModel.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sModelColumn, true));
            this.tbPartNumber.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sPartNumberColumn, true));
            this.tbPrice.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sPriceColumn, true));
            this.tbManufacturer.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sManufacturerColumn, true));
            this.tbManufacturerID.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sManufacturerIDColumn, true));
            this.tbName.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sItemNameColumn, true));
            this.tbType.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sItemTypeColumn, true));
            this.tbCondition.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sConditionColumn, true));
            this.tbSerialNumber.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sSerialNumberColumn, true));
            this.tbLocation.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sLocation, true));
            this.dtpExpiration.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sExpirationColumn, true));
            this.rtbDescription.DataBindings.Add(new System.Windows.Forms.Binding("Text", bsInfo, sItemDescriptionColumn, true));
        }

        private void populateGrid()
        {
            //- Bind the Data Grid to the Binding Source:
            this.dgvInfo.DataSource = bsInfo;
        }

        private void bindNagivation()
        {
            //- Bind the Navigation Tool to the Binding Source:
            this.bnInfo.BindingSource = bsInfo;
        }

        /// <summary>
        /// Purpose: Locates expired items in the Data Grid by looping through each row and
        ///          comparing the date that is in the ExpirationDate column with the 
        ///          current date.  If the value in the ExpirationDate column is less than
        ///          the current date, we add one to our count and move on.  An attempt was 
        ///          made to select the rows that contain expired items, but this feature 
        ///          does not yet work properly in this version.  Further releases will be
        ///          made in the near future.
        /// Known Issues: Currently the selection of rows containing expired items does not
        ///               function properly.  This could be due to the order of events that
        ///               occur upon launching the application.
        /// </summary>
        private void checkForExpiredItems()
        {
            int count = 0;
            string invalidDate = "01/01/0001";  //- Used to check against NULL values.
            DateTime currentDate = DateTime.Now;

            try
            {
                foreach (DataGridViewRow row in dgvInfo.Rows)
                {
                    DateTime expirationDate = Convert.ToDateTime(row.Cells["ExpirationDate"].Value);

                    if (expirationDate < currentDate && expirationDate != Convert.ToDateTime(invalidDate))
                    {
                        dgvInfo.Rows[row.Index].Selected = true;
                        count++;
                    }
                }

                MessageBox.Show("There are " + count + " expired items in the database");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //- Add a new row to the DataTable using a DataRow:
            drRecord = dtProducts.NewRow();
            dtProducts.Rows.Add(drRecord);

            //- Select the new row on the grid:
            bsInfo.Position = dtProducts.Rows.IndexOf(drRecord);
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //- Check for changes in the DataSet and write the changes to the file:
            if (dsInventory.HasChanges())
            {
                dsInventory.GetChanges();
                dsInventory.WriteXml(filePath);
                getDataSource();
            }
            else
            {
                dsInventory.RejectChanges();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //- Close the Application:
            this.Close();
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(userGuide);
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            /* The following method attempts to locate a search term
             * in the DataGridView and select the row if that term is found.
             * Keep in mind that we are searching on the grid itself and not
             * on within the DataTable.  Using this technique, we could
             * capture all of the rows that contain the search term, and 
             * add those DataRows to a second DataTable.  We could then
             * use this DataTable as our filter for the DataGridView. */

            foreach (DataGridViewRow row in dgvInfo.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.FormattedValue.ToString() != String.Empty)
                    {
                        if (cell.Value.ToString() == tbSearch.Text)
                        {
                            dgvInfo.CurrentCell = dgvInfo[cell.ColumnIndex, row.Index];
                            dgvInfo.Rows[row.Index].Selected = true;

                            /* A break was used here to escape the for loop.
                             * This would cause the current row that contained the 
                             * search criteria to be highlighed.  Without placing a 
                             * break at this point, the for loop would continue, and 
                             * cause any rows that were highlighted to not be highlighted. */
                            break;
                        }
                        else
                            cell.Selected = false;
                    }
                    else
                        cell.Selected = false;
                }
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //- Add a new row to the DataTable using a DataRow:
            drRecord = dtProducts.NewRow();
            dtProducts.Rows.Add(drRecord);

            //- Select the new row on the grid:
            bsInfo.Position = dtProducts.Rows.IndexOf(drRecord);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //- Check for changes in the DataSet and write the changes to the file:
            if (dsInventory.HasChanges())
            {
                dsInventory.GetChanges();
                dsInventory.WriteXml(filePath);
                getDataSource();
            }
            else
            {
                dsInventory.RejectChanges();
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            DialogResult dr;

            dr = MessageBox.Show("Are you sure you want to remove this record?", "Remove Record?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                //- If there are rows remaining in the grid:
                if (dgvInfo.SelectedRows.Count > 0)
                {
                    //- Get the selected row from the grid and associate it with a datarow:
                    DataRowView currentDataRowView = (DataRowView)dgvInfo.CurrentRow.DataBoundItem;
                    DataRow drRemove = currentDataRowView.Row;

                    //- Remove the associated datarow from the dataset:
                    dtProducts.Rows.Remove(drRemove);

                    dsInventory.WriteXml(filePath);
                    dsInventory.AcceptChanges();
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            //- Close the Application:
            this.Close();
        }
    }
}
