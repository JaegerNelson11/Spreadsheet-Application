// Jaeger Nelson
// 11789985

using SpreadsheetEngine;
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Spreadsheet_Jaeger_Nelson
{
    public partial class Form1 : Form
    {

        private Spreadsheet spreadsheet;
        public Form1()
        {
            InitializeComponent();
            InitializeDataGrid();

            spreadsheet = new Spreadsheet(50, 26);
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            spreadsheet.CellPropertyChanged += UpdateSpreadsheet;
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            // HW7 Subscribe to the new BeginEdit and EndEdit events
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            dataGridView1.CellBeginEdit += dataGridView1_CellBeginEdit;
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
        }


        public void InitializeDataGrid()
        {
            // Clear existing columns and rows for a fresh grid just in case
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();

            Text = "Jaeger Nelson - 11789985";

            // Loop to add columns A through Z using char var for recursion and our column name
            for (char c = 'A'; c <= 'Z'; c++)
            {
                dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                {
                    HeaderText = c.ToString(),
                    Name = "Column" + c, // Set Column Name and Header Text
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells // Adjust column width automatically for perfect spacing
                });
            }

            // Loop to add rows 1 through 50 using int var for recursion and our row name
            for (int i = 1; i <= 50; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i - 1].HeaderCell.Value = i.ToString();
            }

            // Disable user ability to add/delete rows in case it was set back to true
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;

            // Set size of row header to auto so we can see the row number, initialy was too small.
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;



        }


        private void UpdateSpreadsheet(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "CellChanged")
            {
                // get cell that needs update
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                Cell UpdatedCell = sender as Cell;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                if (UpdatedCell != null)
                {
                    // get row and column
                    int cellRow = UpdatedCell.RowIndex;
                    int cellColumn = UpdatedCell.ColumnIndex;

                    // update that cell in the form
                    dataGridView1.Rows[cellRow].Cells[cellColumn].Value = UpdatedCell.Value;
                }
            }
            else if (e.PropertyName == "CellColorChanged")
            {
                // get the cell that we need to update
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                Cell UpdatedCell = sender as Cell;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                if (UpdatedCell != null)
                {
                    // get its row and column
                    int cellR = UpdatedCell.RowIndex;
                    int cellC = UpdatedCell.ColumnIndex;

                    // get the color from the cell
                    int intColor = (int)UpdatedCell.BGColor;
                    Color color = Color.FromArgb(intColor);

                    // update that cells color
                    dataGridView1.Rows[cellR].Cells[cellC].Style.BackColor = color;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // instantiate a random object
            Random random = new Random();

            // modifying 50 random cells to display a string
            for (int i = 0; i < 50; i++)
            {
                // create a random row 0-25
                // and random column 0-49
                int randomRow = random.Next(0, 50);
                int randomColumn = random.Next(0, 26);

                // set that random cell's value to a string
                Cell current = spreadsheet.GetCell(randomRow, randomColumn);
                current.Text = "Hello World!";
            }

            // modifying every cell in column B
            for (int i = 0; i < 50; i++)
            {
                // set every cell in column B (column 1) to its column name and row number
                Cell current = spreadsheet.GetCell(i, 1);
                current.Text = "This is cell B" + (i + 1);
            }

            // modifying every cell in column B to match column A
            for (int i = 0; i < 50; i++)
            {
                // set every cell in column A (column 0) to column B's value in the same row
                Cell current = spreadsheet.GetCell(i, 0);
                current.Text = "=B" + (i + 1);
            }
        }



        // HW7 code starting here

        // Part 1 Create CellBeginEdit and CellEndEdit events

        // CellBeginEdit
        // Use when the user begins editing a cell

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

            // Get index of row and column of cell to update
            int row = e.RowIndex;
            int column = e.ColumnIndex;

            Cell updateCell = spreadsheet.GetCell(row, column);

            if (updateCell != null)
            {
                // Update the cell to disply its text
                dataGridView1.Rows[row].Cells[column].Value = updateCell.Text;
            }

        }

        // CellEndEdit
        // Use when the user finishes editing a cell

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Get index of row and column of cell to update
            int row = e.RowIndex;
            int column = e.ColumnIndex;

            bool checkEdit = true;


            Cell updateCell = spreadsheet.GetCell(row, column);

            // create a RestoreText command 
            RestoreText[] undoText = new RestoreText[1];


            // store the old text of the cell in case its needed for undo
            string oldText = updateCell.Text;

            // instantiate the RestoreText with the oldText
            undoText[0] = new RestoreText(updateCell, oldText);

            if (updateCell != null)
            {
                // First we will check if the user has deleted the text in the cell
                try
                {
                    // if (the cell's text did not change but there was text in the cell)
                    if (updateCell.Text == dataGridView1.Rows[row].Cells[column].Value.ToString())
                        checkEdit = false;

                    // update the Text property of the cell to notify subscribers
#pragma warning disable CS8601 // Possible null reference assignment.
                    updateCell.Text = dataGridView1.Rows[row].Cells[column].Value.ToString();
#pragma warning restore CS8601 // Possible null reference assignment.
                }
                catch (NullReferenceException)
                {
                    // if the cell didn't have text before and after the edit
                    if (updateCell.Text == null) checkEdit = false;

                    updateCell.Text = "";

                }


                // Update the specific cell to show its value property
                dataGridView1.Rows[row].Cells[column].Value = updateCell.Value;

                // add an undo if the cell was edited
                if (checkEdit == true)
                {
                    // add the text change to the undo stack
                    spreadsheet.AddUndo(new MultiCommand(undoText, "cell text change"));

                    // update the edit menu options to display correctly
                    UpdateMenu();
                }
            }
        }



        // HW 8 Code
        private void chooseBackgroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // instantiate a ColorDialog to pick a color
            ColorDialog colorDialog = new ColorDialog();

            // instaniate a list of Control RestoreColor for multiple color changes
            List<RestoreColor> undoColors = new List<RestoreColor>();

            // if "OK" is selected
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                // loop through all user selected cells
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    // get the cell to update
                    Cell updateCell = spreadsheet.GetCell(cell.RowIndex, cell.ColumnIndex);

                    // save a copy of the cell's old color for undo if needed
                    uint oldCol = updateCell.BGColor;

                    // if the old color was initially 0, make it white
                    if (oldCol == 0) oldCol = (uint)Color.White.ToArgb();

                    // update the background color
                    updateCell.BGColor = (uint)colorDialog.Color.ToArgb();

                    // add the old color to the list of undoColors
                    RestoreColor undoColor = new RestoreColor(updateCell, oldCol);
                    undoColors.Add(undoColor);
                }
            }

            // add all of the color changes to the undo stack
            spreadsheet.AddUndo(new MultiCommand(undoColors.ToArray(), "changing cell background color"));

            // update the edit menu options to display correctly
            UpdateMenu();
        }


        private void UpdateMenu()
        {
            // retrieve the menu items for edit tab
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            ToolStripMenuItem editMenu = menuStrip1.Items[0] as ToolStripMenuItem;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

            // loop through our two options (Undo and Redo)
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            foreach (ToolStripMenuItem menuItem in editMenu.DropDownItems)
            {
                // if it is undo
                if (menuItem.Text.Substring(0, 4) == "Undo")
                {
                    // the undo option being available is dependent on our undo stack
                    menuItem.Enabled = !(spreadsheet.UndoEmpty);

                    // update the menuItem text to display the correctly
                    menuItem.Text = "Undo " + spreadsheet.UndoCmd;
                }
                // check if it's redo
                else if (menuItem.Text.Substring(0, 4) == "Redo")
                {
                    // the redo option being available is dependent on our redo stack
                    menuItem.Enabled = !(spreadsheet.RedoEmpty);

                    // update the menuItem text to display the correct
                    // possible redo command
                    menuItem.Text = "Redo " + spreadsheet.RedoCmd;
                }
            }
#pragma warning restore CS8602 // Dereference of a possibly null reference.



        }

        // Undo has been selected by user, call our functions
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // call the spreadsheet's undo function and update menu
            spreadsheet.Undo();
            UpdateMenu();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // call the spreadsheet's redo function and update menu
            spreadsheet.Redo();
            UpdateMenu();
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // instantiate an OpenFileDialog to help the user load the file
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // if the user clicks OK
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // clear the spreadsheet before we load the file
                ClearSpreadsheet();

                // instantiate the file for loading
                Stream ifile = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read);

                // call our spreadsheet.cs load function
                spreadsheet.Load(ifile);

                // Now we can dispose of the file
                ifile.Dispose();

                // Clear undo and redo stacks
                spreadsheet.ClearUndoRedo();
            }

            // update the menu since we cleared the undo and redo stacks
            UpdateMenu();
        }



        // clears the spreadsheet, used for loading a file
        private void ClearSpreadsheet()
        {
            // loop through every cell in the spreadsheet
            for (int i = 0; i < spreadsheet.RowCount; i++)
            {
                for (int j = 0; j < spreadsheet.ColumnCount; j++)
                {
                    // clear every cell
                    Cell updateCell = spreadsheet.GetCell(i, j);
                    updateCell.Clear();
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // instantiate a SaveFileDialog
            SaveFileDialog saveFile = new SaveFileDialog();

            // when user clicks ok
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                // instantiate an outfile for the save
                Stream ofile = new FileStream(saveFile.FileName, FileMode.Create, FileAccess.Write);

                // call save
                spreadsheet.Save(ofile);

                // dispose of ofile
                ofile.Dispose();
            }

        }
    }
}


