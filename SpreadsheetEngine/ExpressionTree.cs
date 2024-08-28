// Jaeger Nelson
// 11789985

using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    public class ExpressionTree
    {
        private Node root;
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        private static Dictionary<string, double> variables;
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

        // Parent Node Class
        private abstract class Node
        {
            public abstract double Eval();
        }

        // ExpressionTree Constructor
        public ExpressionTree(string expression)
        {
            variables = new Dictionary<string, double>();
            root = Build(expression);
            variables.Clear();
        }

        public void SetVariable(string variableName, double variableValue)
        {
            variables[variableName] = variableValue;
        }

        public string[] GetVariables()
        {
            return variables.Keys.ToArray();
        }


        public double Evaluate()
        {
            if (root == null) 
            {
                return double.NaN;
            }
            else
            {
                return root.Eval();
            }
        }

        private class ConstantNode : Node
        {
            private double value;
            public ConstantNode(double newValue)
            {
                value = newValue;
            }

            // Override Eval from parent based on the evaluation of a constant
            public override double Eval()
            {
                return value;
            }
        }

        private class VariableNode : Node
        {
            private string variableName;
            public VariableNode(string newVariableName) 
            {
                variableName = newVariableName;
            }


            // Override Eval from parent based on the evaluation of a variable. Must be checked with our variable dictionary
            public override double Eval()
            {

                if (!variables.ContainsKey(variableName))
                {
                    variables[variableName] = double.NaN; // Indicate undefined variable
                }

                return variables[variableName];
            }
        }

        private class OperatorNode : Node
        {
            private char op;
            private Node left, right;

            public OperatorNode(char newOp, Node newLeft, Node newRight)
            {
                op = newOp;
                left = newLeft;
                right = newRight;
            }

            // Override Eval from parent based on the evaluation of a binary operator and its left and right nodes
            public override double Eval()
            {
                // Firstly evalue left and right nodes
                double valueLeft = left.Eval();
                double valueRight = right.Eval();

                // Use a switch based on the switch case op variable, the name of the given operator
                switch (op)
                {
                    case '+': return valueLeft + valueRight;
                    case '-': return valueLeft - valueRight;
                    case '*': return valueLeft * valueRight;    
                    case '/': return valueLeft / valueRight;
                    default: return double.NaN;
                }
            }
        }

        private static Node Build(string expression)
        {

            // Clear all spaces form the expression
            expression = expression.Replace(" ", "");

            // Check if expression is completely in parenthesis
            bool isParenth = CheckIfEnclosedParenth(expression);




            // keep checking for enclosed parenthesis
            while (isParenth == true)
            {
                // parse out the enclosing parenthesis
                expression = expression.Substring(1, expression.Length - 2);
                isParenth = CheckIfEnclosedParenth(expression);
            }

            // get the low op index
            int index = GetIndexLowestOpPrec(expression);

            if (index != -1)
            {
                return new OperatorNode(expression[index],
                                   Build(expression.Substring(0, index)),
                                   Build(expression.Substring(index + 1)));
            }


            // the last node will not be an OperatorNode so call function BuildSimple to either assign it as a constantNode or variableNode
            return BuildSimple(expression);
        }

        private static Node BuildSimple(string expression) 
        {
            double value;

            // if the string can be converted to a double
            if (double.TryParse(expression, out value))
            {
                // instantiate and return a new ConstantNode
                return new ConstantNode(value);
            }


            // otherwise instantiate and return a new VariableNode
            return new VariableNode(expression);
        }



        /* HW6 code starts here */

        // Function to get the index of the lowest precedence operator in the given expression.
        // It takes an expression as input and returns the index of the lowest precedence operator.
        // It considers the precedence of operators and also handles parentheses.
        private static int GetIndexLowestOpPrec(string expression)
        {
            int numParenth = 0, index = -1;

            // loop through the full expression
            for (int i = expression.Length - 1; i >= 0; i--)
            {
                switch (expression[i])
                {
                    case ')':
                        numParenth--;
                        break;

                    case '(':
                        numParenth++;
                        break;

                    // if there is  a + or - without anything else inside parenthesis return its index
                    case '+':
                    case '-':
                        if (numParenth == 0) return i;
                        break;

                    // if there is a * or / inside parenthesis return this index
                    case '*':
                    case '/':
                        if (numParenth == 0 && index == -1) index = i;
                        break;
                }
            }
            return index;
        }

        // Function to check if an expression is enclosed within matching parentheses.
        // It takes an expression as input and returns a bool indicating if it is enclosed within parentheses.
        private static bool CheckIfEnclosedParenth(string expression)
        {
            int numParenth = 0;

            // check if the first and last character of the expression are open and closed parenthesis
            if ((expression[0] == '(') && (expression[expression.Length - 1] == ')'))
            {
                // check if they're matching parenthesis
                for (int i = 1; i < expression.Length - 1; i++)
                {
                    switch (expression[i])
                    {
                        case ')':
                            if (numParenth == 0) return false;
                            numParenth--;
                            break;

                        case '(':
                            numParenth++;
                            break;

                        default:
                            break;
                    }
                }

                if (numParenth == 0) return true;
            }

            return false;
        }

    }
}
