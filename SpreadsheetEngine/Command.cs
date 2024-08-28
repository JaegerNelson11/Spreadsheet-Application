using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{

    // HW 8 Code, we will use in main form.cs to implement undos and redos


    public interface Command
    {
        Command Execute();
    }

    public class RestoreColor : Command
    {
        private Cell m_cell;
        private uint m_color;

        public RestoreColor(Cell cell, uint color)
        {
            m_cell = cell;
            m_color = color;
        }

        public Command Execute()
        {
            // build inverse of the cell
            var inv = new RestoreColor(m_cell, m_cell.BGColor);

            // set color
            m_cell.BGColor = m_color;

            // return inverse
            return inv;

        }
    }

    public class RestoreText : Command
    {
        private Cell m_cell;
        private string m_text;

        public RestoreText(Cell cell, string text)
        {
            m_cell = cell;
            m_text = text;
        }

        public Command Execute()
        {
            // build inverse of the cell
            var inverse = new RestoreText(m_cell, m_cell.Text);

            // set text
            m_cell.Text = m_text;

            // return inverse
            return inverse;
        }
    }

    public class MultiCommand
    {
        private Command[] m_commands;
        private string m_comName;

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public MultiCommand() { }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

        public MultiCommand(Command[] commands, string comName)
        {
            m_commands = commands;
            m_comName = comName;
        }

        public string CommandName
        {
            get { return m_comName; }
        }

        public MultiCommand Execute()
        {
            // initialize a list of commands
            List<Command> comList = new List<Command>();

            // loop through the member array of commands
            foreach (Command cmd in m_commands)
            {
                // add the command to the list
                comList.Add(cmd.Execute());
            }

            // return the inverse array
            return new MultiCommand(comList.ToArray(), m_comName);
        }
    }
}
