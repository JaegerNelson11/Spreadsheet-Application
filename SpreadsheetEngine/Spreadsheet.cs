// Jaeger Nelson
// 11789985
using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Formats.Asn1;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading.Channels;
using System.Xml;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;



namespace SpreadsheetEngine
{

    public class Spreadsheet
    {
        private int m_rows;
        private int m_columns;
        private Cell[,] m_spreadsheet;

        public event PropertyChangedEventHandler CellPropertyChanged = delegate { };

        private Dictionary<Cell, List<Cell>> dependencies; // HW 7 This is where we will note dependencies between cells refrencing cells

        // HW8 new members
        private Stack<MultiCommand> m_undos = new Stack<MultiCommand>();
        private Stack<MultiCommand> m_redos = new Stack<MultiCommand>();

        private HashSet<Cell> m_visitedCells = new HashSet<Cell>();

       
        //Create new instance of a cell since it is an abstract class
        private class BasicCell : Cell
        {
            public BasicCell(int newRow, int newColumn) : base(newRow, newColumn) { }
        }

        public Spreadsheet(int newRows, int newColumns)
        {
            m_rows = newRows;
            m_columns = newColumns;
            dependencies = new Dictionary<Cell, List<Cell>>();

            // initialize the array of cells that is the spreadsheet
            m_spreadsheet = new Cell[newRows, newColumns];

            // cycle through each cell
            for (int i = 0; i < newRows; i++)
            {
                for (int j = 0; j < newColumns; j++)
                {
                    // instantiate a cell
                    Cell newCell = new BasicCell(i, j);

                    // subscribe the spreadsheet to every cell
#pragma warning disable CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).
                    newCell.PropertyChanged += UpdateCell;
#pragma warning restore CS8622 // Nullability of reference types in type of parameter doesn't match the target delegate (possibly because of nullability attributes).

                    // add the cell into the spreadsheet in its corresponding position
                    m_spreadsheet[i, j] = newCell;
                }
            }
        }

        // Getters for rowcount, columncount, and for a specific cell
        public int ColumnCount
        {
            get { return m_columns; }
        }
        public int RowCount
        {
            get { return m_rows; }
        }
        public Cell GetCell(int row, int column)
        {
            return m_spreadsheet[row, column];
        }




        // HW7 Code

        // overloaded GetCell that we can use when we have a cells name
        public Cell GetCell(string cellName)
        {
            // check if the first character is a letter
            if (!Char.IsLetter(cellName[0]))
            {
                // the first character was not a letter return null
#pragma warning disable CS8603 // Possible null reference return.
                return null;
#pragma warning restore CS8603 // Possible null reference return.
            }

            // check if the first letter is capitalized
            if (!Char.IsUpper(cellName[0]))
            {
                // the first character is a letter, but it isn't upper case return null
#pragma warning disable CS8603 // Possible null reference return.
                return null;
#pragma warning restore CS8603 // Possible null reference return.
            }

            int tempCellColumn = cellName[0] - 'A';
            int tempCellRow;
            Cell tempCell;

            // parse the rest cell name to get row
            // check for a valid int row
            if (!Int32.TryParse(cellName.Substring(1), out tempCellRow))
            {
                // the rest of the string was not an int return null
#pragma warning disable CS8603 // Possible null reference return.
                return null;
#pragma warning restore CS8603 // Possible null reference return.
            }

            // try to retrieve cell now
            try
            {
                // subtract one because array starts at 0,0
                tempCell = GetCell(tempCellRow - 1, tempCellColumn);
            }
            catch (Exception)
            {
                // retrieving the cell was not possible due to the row not existing
#pragma warning disable CS8603 // Possible null reference return.
                return null;
#pragma warning restore CS8603 // Possible null reference return.
            }

            // return the cell as all checks passed at this point
            return tempCell;
        }

        // Removes the dependencies for a cell
        private void RemoveDependencies(Cell currentCell)
        {
            foreach (List<Cell> cells in dependencies.Values)
            {
                if (cells.Contains(currentCell))
                {
                    cells.Remove(currentCell);
                }
            }
        }

        // Adds dependencies for a cell
        private void AddDependencies(Cell currentCell, string[] independents)
        {
            foreach (string independent in independents)
            {
                Cell independentC = GetCell(independent);

                // if the dictionary does not contain the key
                if (!dependencies.ContainsKey(independentC))
                {
                    // create a new list if the cell is a new key
                    dependencies[independentC] = new List<Cell>();
                }

                // add the referenced cell to the references
                dependencies[independentC].Add(currentCell);
            }
        }

        // Updates all dependent cells
        private void UpdateDependencies(Cell currentCell)
        {
            foreach (Cell dependent in dependencies[currentCell].ToArray())
            {
                UpdateCell(dependent);
            }
        }

        // this function recieves a certain event argument e, and changes the cell based on the argument given
        private void UpdateCell(object sender, PropertyChangedEventArgs e)
        {
            // if the text property is being changed, only change the cell's text 
            if (e.PropertyName == "Text")
            {
                // get the cell that we need to update
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                Cell UpdatedCell = sender as Cell;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.


#pragma warning disable CS8604 // Possible null reference argument.
                UpdateCell(UpdatedCell);
#pragma warning restore CS8604 // Possible null reference argument.

            }
            else if (e.PropertyName == "Color")
            {
                // get the cell in need of update
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                Cell cellToUpdate = sender as Cell;
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                // and notify subscribers that the cell property changed
                CellPropertyChanged(cellToUpdate, new PropertyChangedEventArgs("CellColorChanged"));
            }
        }

        // Overloaded UpdateCell that takes cell as parameter
        // This is where HW5 and 6 Expression tree have been implemented
        private void UpdateCell(Cell updateCell)
        {
            // Part 3, clear old dependencies
            RemoveDependencies(updateCell);

            // first check if the cell is empty as in the user deleted the text from the cell
            if (string.IsNullOrEmpty(updateCell.Text))
            {
                updateCell.Value = "";
            }

            // Then an else if to see if the text of the cell is not an equation
            else if (updateCell.Text[0] != '=')
            {
                double value;

                // Check if the cell contains a double
                if (double.TryParse(updateCell.Text, out value))
                {
                    // If so, then build a new expression tree for the new variable
                    ExpressionTree newTree = new ExpressionTree(updateCell.Text);
                    value = newTree.Evaluate();
                    newTree.SetVariable(updateCell.Name, value);


                    // update the value to the double
                    updateCell.Value = value.ToString();
                }
                else
                {
                    // Otherwise just set to the text
                    updateCell.Value = updateCell.Text;
                }
            }

            // The cell contains an expression that we need to evaluate
            else
            {


                // New bool variable to check for errors
                bool error = false;

                // parse the text to get the expression
                string expression = updateCell.Text.Substring(1);


                // Now we can build the expression tree
                ExpressionTree newTree = new ExpressionTree(expression);
                newTree.Evaluate();


                // Get the varaible names dictionary
                string[] variables = newTree.GetVariables();

                // Loop through each element of the variables array
                foreach (string variable in variables)
                {
                    // get the cell and add its value to the dictionary
                    double value = double.NaN;
                    Cell currentCell = GetCell(variable);

                    // HW10: checks for valid cells
                    // Case 1: BAD REFERENCE
                    if (currentCell == null)
                    {
                        // display cell has a bad reference and notify subscribers
                        updateCell.Value = "!(bad reference)";
                        CellPropertyChanged(updateCell, new PropertyChangedEventArgs("CellChanged"));

                        // set error to true
                        error = true;
                    }

                    // Case 2: SELF-REFERENCE
                    else if (updateCell.Name == variable)
                    {
                        // display cell has a self-reference and notify subscribers
                        updateCell.Value = "!(self reference)";
                        CellPropertyChanged(updateCell, new PropertyChangedEventArgs("CellChanged"));

                        // set error to true
                        error = true;
                    }

                    // if there's an error, we want to step out of the function
                    if (error == true)
                    {
                        // if cells depend on the one we just updated
                        if (dependencies.ContainsKey(updateCell))
                        {
                            // update all dependent cells
                            UpdateDependencies(updateCell);
                        }

                        return;
                    }


                    // try to parse out the double
                    if (currentCell != null )
                    double.TryParse(currentCell.Value, out value);

                    // set the variable
                    newTree.SetVariable(variable, value);

                    // add the cell to the HashSet to check for circular references
#pragma warning disable CS8604 // Possible null reference argument.
                    m_visitedCells.Add(currentCell);
#pragma warning restore CS8604 // Possible null reference argument.

                }

                // set the value to the computed value of the ExpTree
                updateCell.Value = newTree.Evaluate().ToString();

                // add the dependencies
                AddDependencies(updateCell, variables);

                // Case 3: CIRCULAR REFRENCE
                // do a check for circular reference
                if (CheckCircularReferences(updateCell))
                {
                    // display cell has circular-reference and notify subscribers
                    updateCell.Value = "!(circular reference)";
                    CellPropertyChanged(updateCell, new PropertyChangedEventArgs("CellChanged"));

                    // set error to true
                    error = true;
                }

                // if there's an error, return
                if (error == true) return;
            }



            // if there are cells that depend on the updated one
            if (dependencies.ContainsKey(updateCell))
            {
                // update all dependent cells
                UpdateDependencies(updateCell);
            }

            // and notify subscribers that the cell property changed
            CellPropertyChanged(updateCell, new PropertyChangedEventArgs("CellChanged"));

            m_visitedCells.Clear();
        }


        // HW 10 code for undo circular reference

        public bool CheckCircularReferences(Cell currentCell)
        {
            // if we can't add the cell to the hash set, there's a circular reference
            if (m_visitedCells.Add(currentCell) == false) return true;
            return false;
        }



        // HW 8 code for undo redo

        // UndoEmpty property
        // checks if there's something on the member undo stack
        public bool UndoEmpty
        {
            get { return m_undos.Count == 0; }
        }

        // RedoEmpty property
        // checks if there's something on the redo stack
        public bool RedoEmpty
        {
            get { return m_redos.Count == 0; }
        }

        // UndoCommand property
        public string UndoCmd
        {
            get
            {
                // check to make sure undo stack is not empty with our helper function
                if (!UndoEmpty)
                {
                    // return the command name
                    return m_undos.Peek().CommandName;
                }

                // otherwise return an empty string if there's nothing to undo
                return "";
            }
        }

        // RedoCommand property
        public string RedoCmd
        {
            get
            {
                // check to make sure redo stack is not empty with our helper function
                if (!RedoEmpty)
                {
                    // return the command name
                    return m_redos.Peek().CommandName;
                }

                // otherwise return an empty string if there's
                // nothing to redo
                return "";
            }
        }

        // adds commands to the undo stack
        public void AddUndo(MultiCommand undos)
        {
            // push the undos onto the undo stack
            m_undos.Push(undos);
        }


        // performs an undo
        public void Undo()
        {
            // check if the undo stack is empty
            if (!UndoEmpty)
            {
                // pop commands off the undo stack
                MultiCommand commands = m_undos.Pop();

                // push the commands onto the redo stack
                m_redos.Push(commands.Execute());
            }
        }

        // performs a redo
        public void Redo()
        {
            // check if the redo stack is empty
            if (!RedoEmpty)
            {
                // pop commands off the redo stack
                MultiCommand commands = m_redos.Pop();

                // push the commands onto the undo stack
                m_undos.Push(commands.Execute());
            }
        }

        
        // HW 9 Code Load and Save features in spreadsheet class

        public void Load (Stream ifile)
        {
            // load the xml file from our stream parameter
            XDocument xml = XDocument.Load(ifile);

            // loop through all cell tags in our XML file
#pragma warning disable CS8602 // Dereference of a possibly null reference.
            foreach (XElement tag in xml.Root.Elements("cell"))
#pragma warning restore CS8602 // Dereference of a possibly null reference.
            {
                // Put the cell in memory
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                string cellName = tag.Attribute("name").Value;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                Cell xmlCell = GetCell(cellName);

                // Load all of the cells properties based on the properties in the xml if the xml property is valid
                // Check if xml text property is valid
                if (tag.Element("text") != null)
                {
                    // Load cell text
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                    xmlCell.Text = tag.Element("text").Value.ToString();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                }

                // Get the background color element
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                XElement bgColorElement = tag.Element("bgcolor");
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

                // Check if the bgcolor element exists and is not empty
                if (bgColorElement != null && !string.IsNullOrEmpty(bgColorElement.Value))
                {
                    string hexColor = bgColorElement.Value;
                    if (hexColor.Length == 6) // For "RRGGBB" format
                    {
                        hexColor = "FF" + hexColor; // Prefix with FF (full opacity)
                    }
                    // Convert the string to a uint for the background color
                    uint xmlColor = Convert.ToUInt32(hexColor, 16); // 16 for a hexadecimal base
                    // Load the cell's background color
                    xmlCell.BGColor = xmlColor;
                }

            }
        }

        public void Save(Stream ofile)
        {
            // Create an XmlWriterSettings object to configure the XmlWriter and how it will write
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true; // will format the XML with indentation

            // Create an XmlWriter
            using (XmlWriter xmlFile = XmlWriter.Create(ofile, settings))
            {
                // Write the spreadsheet start tag
                xmlFile.WriteStartElement("spreadsheet");

                // Loop through every cell in the spreadsheet
                for (int i = 0; i < m_rows; i++)
                {
                    for (int j = 0; j < m_columns; j++)
                    {
                        Cell tempCell = m_spreadsheet[i, j];

                        // Only save a cell if it has been modified by the user
                        if (!string.IsNullOrEmpty(tempCell.Text) || tempCell.BGColor != 0)
                        {
                            // Write the start tag for the cell
                            xmlFile.WriteStartElement("cell");

                            // Add the name attribute to the cell element
                            xmlFile.WriteAttributeString("name", tempCell.Name);

                            // Write the bgcolor element
                            // Convert the integer color value to a hex string with 8 characters 
                            string hexColor = tempCell.BGColor.ToString("X8");
                            xmlFile.WriteElementString("bgcolor", hexColor);

                            // Write the text element
                            xmlFile.WriteElementString("text", tempCell.Text);

                            // Close the cell tag
                            xmlFile.WriteEndElement();
                        }
                    }
                }

                // Close the spreadsheet tag
                xmlFile.WriteEndElement();

                // The XmlWriter automatically closes when exiting the using block
            }
        }

        // clears the undo and redo stacks, used for loading a file
        public void ClearUndoRedo()
        {
            m_undos.Clear();
            m_redos.Clear();
        }


    }
}


