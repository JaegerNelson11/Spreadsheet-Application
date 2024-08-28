// Jaeger Nelson
// 11789985

using System;
using System.ComponentModel;
using System.Data.Common;
using System.Formats.Asn1;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SpreadsheetEngine
{
    // Inherit from INotifyPropertyChanged to get access to the event property handler we will use to modify cell values
    public abstract class Cell : INotifyPropertyChanged
    {
        protected int rowI;
        protected int columnI;
        protected string m_text;
        protected string m_value;
        protected string m_name;
        protected uint m_color;

#pragma warning disable CS8612 // Nullability of reference types in type doesn't match implicitly implemented member.
        public event PropertyChangedEventHandler PropertyChanged = delegate { };
#pragma warning restore CS8612 // Nullability of reference types in type doesn't match implicitly implemented member.

        // Default cell empty constructor
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public Cell()
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        {

        }

        // Cell constructor with both index var
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public Cell(int newRowI, int newColumnI)
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        {
            rowI = newRowI;
            columnI = newColumnI;
            m_text = "";
            m_value = "";
            m_color = 0;

            // Convert Index Values to a cell name
            m_name += Convert.ToChar('A' + newColumnI);        // ascii 'A' + column = corresponding letter
            m_name += (newRowI + 1).ToString();                // rows start at one in the user spreadsheet so +1

        }

        // ColumnIndex and RowIndex are both to be read only represented here
        public int ColumnIndex
        {
            get { return columnI; }
        }

        public int RowIndex
        {
            get { return rowI; }
        }

        // Text Property for cell
        public string Text
        {
            get { return m_text; }

            set
            {
                // if the text is being changed to the same text then ignore it
                if (m_text == value) return;

                // otherwise update m_text
                m_text = value;

                // and notify subscribers that the property changed
                PropertyChanged(this, new PropertyChangedEventArgs("Text"));
            }
        }

        // Value property for cell
        public string Value
        {
            get { return m_value; }

            // protect the value from being set by anything other than the spreadsheet class using a protected internal set to make sure
            protected internal set
            {
                // if the value is being changed to the same value then ignore it
                if (m_value == value) return;

                // otherwise update m_value
                m_value = value;

                // and notify subscribers that the property changed
                PropertyChanged(this, new PropertyChangedEventArgs("Value"));
            }
        }

        // New property for Name
        public string Name
        {
            get { return m_name; }
        }


        // Property for cell background color m_color
        public uint BGColor
        {
            get { return m_color; }

            set
            {
                // if the color is being changed to the same color then ignore it
                if (m_color == value) return;

                // otherwise update the color
                m_color = value;

                // and notify subscribers that the property changed
                PropertyChanged(this, new PropertyChangedEventArgs("Color"));
            }
        }

        // resets the text, value and color back to their default values
        public void Clear()
        {
            Text = "";
            Value = "";
            BGColor = 0;
        }
    }
}



    

