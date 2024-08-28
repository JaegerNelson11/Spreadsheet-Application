using Spreadsheet_Jaeger_Nelson;
using SpreadsheetEngine;
using System.ComponentModel;
using System.Text;


// Jaeger Nelson
// 11789985
namespace Spreadsheet_Jaeger_Nelson_Tests
{
    public class Tests
    {
        [TestFixture]
        public class DataGridViewTests
        {
            public Form1 testForm;

            [SetUp]
            public void Setup()
            {
                // Initialize the form
                testForm = new Form1();
            }

            [Test]
            public void TestInitializeDataGrid_ColumnsCount()
            {


                // Test initializedatagrid function and test for amount of columns
                testForm.InitializeDataGrid();
                int actualC = 26;

                // Assert
                Assert.That(actualC, Is.EqualTo(testForm.dataGridView1.ColumnCount));
            }

            [Test]
            public void TestInitializeDataGrid_RowsCount()
            {
                // Test initializedatagrid function and test for amount of rows
                testForm.InitializeDataGrid();
                int actualR = 50;

                // Assert
                Assert.That(actualR, Is.EqualTo(testForm.dataGridView1.RowCount));
            }


            // HW7 Tests

            [Test]
            public void EmptyCellShouldSetEmptyValue()
            {
                var spreadsheet = new Spreadsheet(1, 1);
                var cell = spreadsheet.GetCell(0, 0);

                cell.Text = "";

                Assert.That(cell.Value, Is.EqualTo(""));
            }

            [Test]
            public void CellNumericalTextSetCellValueToText()
            {
                var spreadsheet = new Spreadsheet(1, 1);
                var cell = spreadsheet.GetCell(0, 0);

                cell.Text = "123";

                Assert.That(cell.Value, Is.EqualTo("123"));
            }

            [Test]
            public void TextCellSetCellValueToText()
            {
                var spreadsheet = new Spreadsheet(1, 1);
                var cell = spreadsheet.GetCell(0, 0);

                cell.Text = "Hello";

                Assert.That(cell.Value, Is.EqualTo("Hello"));
            }

            [Test]
            public void FormulaCellEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "5";
                cellB1.Text = "=A1*2";

                Assert.That(cellB1.Value, Is.EqualTo("10"));
            }

            [Test]
            public void AdditionFormulaEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "5";
                cellB1.Text = "=A1 + 3";

                Assert.That(cellB1.Value, Is.EqualTo("8")); 
            }

            [Test]
            public void SubtractionFormulaEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "10";
                cellB1.Text = "=A1 - 5"; 

                Assert.That(cellB1.Value, Is.EqualTo("5")); 
            }

            [Test]
            public void MultiplicationFormulaEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "6";
                cellB1.Text = "=A1 * 2";

                Assert.That(cellB1.Value, Is.EqualTo("12"));
            }

            [Test]
            public void DivisionFormulaEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "20";
                cellB1.Text = "=A1 / 4"; 

                Assert.That(cellB1.Value, Is.EqualTo("5")); 
            }

            [Test]
            public void ReferencingOtherCellEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "10";
                cellB1.Text = "=A1";

                Assert.That(cellB1.Value, Is.EqualTo("10")); 
            }

            [Test]
            public void OrderOfOperationsWithParenthesesEvaluateFormulaAndSetValue()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "2";
                cellB1.Text = "=(3 + 4) * (5 - A1)"; 

                Assert.That(cellB1.Value, Is.EqualTo("21")); 
            }

            [Test]
            public void CellReferencingEmptyCellEvaluateTo0()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellB1 = spreadsheet.GetCell(0, 1);
                cellA1.Text = "=B1";
                

                Assert.That(cellA1.Value, Is.EqualTo("0")); 
            }

            [Test]
            public void ChangeCellColor_ShouldUpdateColorCorrectly()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                uint newColor = 0xFF0000; 

                cellA1.BGColor = newColor;

                Assert.That(cellA1.BGColor, Is.EqualTo(newColor)); 
            }

        }
        [TestFixture]
        public class CommandTestsForUndoRedo
        {
            [Test]
            public void RestoreColorSetCellBackgroundColor()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                RestoreColor command = new RestoreColor(cellA1, 0xFF0000);

                command.Execute();

                Assert.That(cellA1.BGColor, Is.EqualTo(0xFF0000));
            }
            [Test]
            public void RestoreTextSetTextToCell()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var command = new RestoreText(cellA1, "Hello World");

                command.Execute();

                Assert.That(cellA1.Text, Is.EqualTo("Hello World"));
            }

            [Test]
            public void MultiCommandExecuteAllCommands()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                Command[] commands =
                {
                new RestoreColor(cellA1, 0xFF0000),
                new RestoreText(cellA1, "Hello World")
                };
                var multiCommand = new MultiCommand(commands, "TestCommand");

                multiCommand.Execute();

                Assert.That(cellA1.BGColor, Is.EqualTo(0xFF0000));
                Assert.That(cellA1.Text, Is.EqualTo("Hello World"));
            }

            [Test]
            public void UndoWithNoCommandsDoesNotThrowException()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                Assert.DoesNotThrow(() => spreadsheet.Undo());
            }

            [Test]
            public void RedoWithNoCommandsDoesNotThrowException()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                Assert.DoesNotThrow(() => spreadsheet.Redo());
            }
        }

        // HW 9 Tests
        [TestFixture]
        public class SaveAndLoadTests
        {
            // Test case for loading an empty XML file
            [Test]
            public void LoadEmptyFileEmptySpreadsheet()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Define an empty XML string
                string emptyXml = "<spreadsheet></spreadsheet>";

                // Create a stream that we can pass to Load() from the XML string
                using (Stream testStream = new MemoryStream(Encoding.UTF8.GetBytes(emptyXml)))
                {
                    // Pass the new xml file to load
                    spreadsheet.Load(testStream);

                    // Test to see if spreadsheet is empty
                    Assert.That(spreadsheet.GetCell(0, 0).Text, Is.EqualTo(string.Empty));
                    Assert.That(spreadsheet.GetCell(0, 0).BGColor, Is.EqualTo(0));
                }
            }

            // Test case for correctly loading an XML file that contains info for a single cell
            [Test]
            public void LoadSingleCellCorrectlyLoaded()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Define an XML string with a single cell
                string xml = "<spreadsheet>"
                           + "  <cell name='A1'>"
                           + "    <bgcolor>FFFF8000</bgcolor>"
                           + "    <text>=20</text>"
                           + "  </cell>"
                           + "</spreadsheet>";

                // Create a stream from the XML string
                using (Stream testStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
                {
                    // Pass the new xml file to load
                    spreadsheet.Load(testStream);

                    // Verify the cell is correctly loaded and its properties are correct
                    var cell = spreadsheet.GetCell(0, 0);
                    Assert.That(cell.Text, Is.EqualTo("=20"));
                    Assert.That(cell.Value, Is.EqualTo("20"));
                    Assert.That(cell.BGColor, Is.EqualTo(0xFFFF8000));
                }
            }

            // Test case for correctly loading an XML file that contains info for multiple cells
            [Test]
            public void LoadMultipleCellsCorrectlyLoaded()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Define an XML string with multiple cells
                string xml = "<spreadsheet>"
                           + "  <cell name='A1'>"
                           + "    <bgcolor>FFFF0000</bgcolor>"
                           + "    <text>=15</text>"
                           + "  </cell>"
                           + "  <cell name='B2'>"
                           + "    <bgcolor>FF00FF00</bgcolor>"
                           + "    <text>=A1+15</text>"
                           + "  </cell>"
                           + "</spreadsheet>";

                // Create a stream from the XML string
                using (Stream testStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
                {
                    // Load the XML file
                    spreadsheet.Load(testStream);

                    // Verify each cell is correctly loaded and their properties are correct
                    var cellA1 = spreadsheet.GetCell(0, 0);
                    Assert.That(cellA1.Text, Is.EqualTo("=15")); // Text Should be "=15". Value should be int 15
                    Assert.That(cellA1.Value, Is.EqualTo("15"));
                    Assert.That(cellA1.BGColor, Is.EqualTo(0xFFFF0000));

                    var cellB2 = spreadsheet.GetCell(1, 1);
                    Assert.That(cellB2.Text, Is.EqualTo("=A1+15")); // Text Should be "=A1+15", but its value should be evaluated and should be int 30
                    Assert.That(cellB2.Value, Is.EqualTo("30"));
                    Assert.That(cellB2.BGColor, Is.EqualTo(0xFF00FF00));
                }
            }
            



            // Test case for correctly loading an XML file that contains info for multiple color formats for multiple cells
            [Test]
            public void LoadDifferentFormatColors()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Define an XML string with cells containing different color formats to see if the load function can handle them
                string xml = "<spreadsheet>"
                           + "  <cell name='A1'>"
                           + "    <bgcolor>FF800000</bgcolor>"
                           + "    <text>=20</text>"
                           + "  </cell>"
                           + "  <cell name='B1'>"
                           + "    <bgcolor>800000FF</bgcolor>"
                           + "    <text>=40</text>"
                           + "  </cell>"
                           + "</spreadsheet>";

                // Create a stream from the XML string
                using (Stream testStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
                {
                    // Load the XML file
                    spreadsheet.Load(testStream);

                    // Verify each cell is correctly loaded
                    var cellA1 = spreadsheet.GetCell(0, 0);
                    Assert.That(cellA1.Text, Is.EqualTo("=20"));
                    Assert.That(cellA1.BGColor, Is.EqualTo(0xFF800000));

                    var cellB1 = spreadsheet.GetCell(0, 1);
                    Assert.That(cellB1.Text, Is.EqualTo("=40"));
                    Assert.That(cellB1.BGColor, Is.EqualTo(0x800000FF));
                }
            }

            // Test case for correctly saving an empty spreadsheet to an XML file
            [Test]
            public void SaveEmptySpreadsheet_ProducesExpectedXML()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Save empty spreadsheet to a stream
                using (var testStream = new MemoryStream())
                {
                    spreadsheet.Save(testStream);

                    // Convert the memory stream to a string
                    testStream.Position = 0;
                    string xmlContent = Encoding.UTF8.GetString(testStream.ToArray());
                    xmlContent = xmlContent.Substring(1);

                    // Expected XML with XML declaration and an empty spreadsheet
                    string expectedXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"
                                      + "<spreadsheet />";

                    // Compare the saved XML with the expected XML
                    Assert.That(xmlContent, Is.EqualTo(expectedXml));
                }
            }

            // Test case for correctly saving a spreadsheet with a single modified cell to an XML file
            [Test]
            public void SaveSingleCellCorrectlySaved()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Set up a single cell with text and a background color
                var cell = spreadsheet.GetCell(0, 0);
                cell.Text = "=20";
                cell.BGColor = 0xFFFFFF00;

                // Save the spreadsheet to a stream
                using (var testStream = new MemoryStream())
                {
                    spreadsheet.Save(testStream);

                    // Convert the memory stream to a string
                    testStream.Position = 0;
                    string xmlContent = Encoding.UTF8.GetString(testStream.ToArray());
                    xmlContent = xmlContent.Substring(1); // To ignore invisible Byte Order Mark BOM
                                                          // Expected XML content for a single cell
                    string expectedXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"
                                       + "<spreadsheet>\r\n"
                                       + "  <cell name=\"A1\">\r\n"
                                       + "    <bgcolor>FFFFFF00</bgcolor>\r\n"
                                       + "    <text>=20</text>\r\n"
                                       + "  </cell>\r\n"
                                       + "</spreadsheet>";

                    // Compare the saved XML with the expected XML
                    Assert.That(xmlContent, Is.EqualTo(expectedXml));
                }
            }

            // Test case for correctly saving a spreadsheet with multiple modified cells to an XML file
            [Test]
            public void SaveMultipleCells_CorrectlySavesToXML()
            {
                var spreadsheet = new Spreadsheet(2, 2);

                // Set up cell A1 with text and a background color
                var cellA1 = spreadsheet.GetCell(0, 0);
                cellA1.Text = "=20";
                cellA1.BGColor = 0xFFFFFF00;

                // Set up cell B2 with text and a background color
                var cellB2 = spreadsheet.GetCell(1, 1);
                cellB2.Text = "=A1+20";
                cellB2.BGColor = 0xFFFFFF00;

                // Save the spreadsheet to a stream
                using (var testStream = new MemoryStream())
                {
                    spreadsheet.Save(testStream);

                    // Convert the memory stream to a string
                    testStream.Position = 0;
                    string xmlContent = Encoding.UTF8.GetString(testStream.ToArray());
                    xmlContent = xmlContent.Substring(1);


                    // Expected XML content for multiple cells (A1 and B2)
                    string expectedXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"
                                      + "<spreadsheet>\r\n"
                                      + "  <cell name=\"A1\">\r\n"
                                      + "    <bgcolor>FFFFFF00</bgcolor>\r\n"
                                      + "    <text>=20</text>\r\n"
                                      + "  </cell>\r\n"
                                      + "  <cell name=\"B2\">\r\n"
                                      + "    <bgcolor>FFFFFF00</bgcolor>\r\n"
                                      + "    <text>=A1+20</text>\r\n"
                                      + "  </cell>\r\n"
                                      + "</spreadsheet>";

                    // Compare the saved XML with the expected XML
                    Assert.That(xmlContent, Is.EqualTo(expectedXml));
                }
            }



        }

        [TestFixture]
        public class RefrenceErrorTests
        {
            [Test]
            public void TestCellWithBadReference()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);

                cellA1.Text = "=CC12"; // Referencing a non-existent cell

                Assert.That(cellA1.Value, Is.EqualTo("!(bad reference)"));
            }

            [Test]
            public void TestCellWithBadReference2()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);

                cellA1.Text = "=a10"; // Referencing a non-existent cell

                Assert.That(cellA1.Value, Is.EqualTo("!(bad reference)"));
            }

            [Test]
            public void TestCellWithBadReference3()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);

                cellA1.Text = "=C105"; // Referencing a non-existent cell

                Assert.That(cellA1.Value, Is.EqualTo("!(bad reference)"));
            }

            [Test]
            public void TestCellWithSelfReference()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);

                cellA1.Text = "=A1"; // Referencing itself

                Assert.That(cellA1.Value, Is.EqualTo("!(self reference)"));
            }

            [Test]
            public void TestCellsWithCircularRefrence()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellA2 = spreadsheet.GetCell(0, 1);
                var cellB1 = spreadsheet.GetCell(1, 0);
                var cellB2 = spreadsheet.GetCell(1, 1);

                cellA1.Text = "=B1*2"; // Building a circular refrence (one in HW10 instructions)
                cellB1.Text = "=B2*3";
                cellB2.Text = "=A2*4";
                cellA2.Text = "=A1*5";

                Assert.That(cellA1.Value, Is.EqualTo("!(circular reference)"));
            }

            [Test]
            public void TestCellsWithCircularRefrence2()
            {
                var spreadsheet = new Spreadsheet(2, 2);
                var cellA1 = spreadsheet.GetCell(0, 0);
                var cellA2 = spreadsheet.GetCell(0, 1);
                var cellB1 = spreadsheet.GetCell(1, 0);
                var cellB2 = spreadsheet.GetCell(1, 1);

                cellA1.Text = "=B1*7"; // Building a circular refrence (one in video demo)
                cellB1.Text = "=B2";
                cellB2.Text = "=A2";
                cellA2.Text = "=A1";

                Assert.That(cellA1.Value, Is.EqualTo("!(circular reference)"));
            }
        }                         
    }
}

