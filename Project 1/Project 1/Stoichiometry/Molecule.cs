/* Name: Katherine Haldane
 * Date: February 13, 2014
 * Description: Library assembly using C# that provides stoichiometric services and a 
 *               windows client to demonstrate the library. Allows the user to key-in molecular formula 
 *               for chemica compound. Validates formula to encure no illegal symbols and subscripts. 
 *               Displays molecular weight of a molecule if valid, normalize the formula, and save the
 *               calculated formula to the database.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Forms;

namespace Stoichiometry
{
    //Public interface containing Formula, formula, Valid, MolecularWeight, Normalize, Save, and FormulaList
    public interface IMolecule
    {
        string Formula { get; set; }
        string formula { get; }
        bool Valid();
        double MolecularWeight { get; }
        void Normalize();
        bool Save();
        string[] FormulasList { get; }
    }


    public class Molecule : IMolecule
    {
        //Global Scope
        Dictionary<string, int> normDict = new Dictionary<string, int>();
        Dictionary<string, double> eleDict = new Dictionary<string, double>();
        List<string> formList = new List<string>();
        List<string> regexElements = new List<string>();
        List<double> calcTotals = new List<double>();
        string elements;
        double newTotal;

        // Data-related member variables 
        private DataSet ds;
        private OleDbDataAdapter odba;


        public string Formula { get; set; }
        //formula gets Formula (due to macro name in Excel spreadsheet)
        public string formula
        {
            get
            {
                return Formula;
            }
        }

        //Valid function that checks the Formula to make sure it does not have invalid characers, whitespaces, the element exists, etc.
        public bool Valid()
        {
            try
            {
                InitializeElement();
                string str = Formula;

                //Add each element from the database to a list
                foreach (DataRow row in ds.Tables["Elements"].Rows)
                {
                    string ele = row.Field<string>("Symbol");
                    regexElements.Add(ele); //Add element to List
                }


                if (Formula == "")
                {
                    return true;
                }
                else
                {
                    //Check Formula to make sure it meets all validations
                    bool match = Regex.IsMatch(Formula, "^([A-Z]{1}[a-z]*([1-9]{1}[0-9]*){0,1})+$");
                    if (match == true)
                    {
                        //Go through database and add each symbol to string elements
                        foreach (string symbol in regexElements)
                        {
                            elements += "(^" + symbol + "{1}$)|";
                        }

                        //Remove the last element in the string elements because it will be |
                        elements = elements.Remove(elements.Length - 1);

                        //Add ONLY elements to a word array
                        MatchCollection matches = Regex.Matches(Formula, @"([A-Z][a-z]*)");
                        string[] elementLetter = matches.Cast<Match>().Select(m => m.Value).ToArray();

                        //Compare all the strings in array to the element list
                        for (int i = 0; i < elementLetter.Count(); i++)
                        {
                            bool regexMatch = Regex.IsMatch(elementLetter[i], @"^(" + elements + ")+$");

                            //If match does not exist return false else return true
                            if (regexMatch == false)
                            {
                                return false;
                            }
                        }
                        return true;
                    }
                    else
                    {
                        //Return false Formula is not valid
                        if (Formula == "")
                        {
                            return true;
                        }
                        else
                            return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (Formula == "")
                {
                    return true;
                }
                else
                    return false;
            }
        }

        //Calculate the molecular weight of Formula
        public double MolecularWeight
        {
            get
            {
                //Check if Formula is valid
                if (Valid() == true)
                {
                    try
                    {
                        //Clear element dictionary and calculation totals
                        eleDict.Clear();
                        calcTotals.Clear();
                        //Open database
                        InitializeElement();
                        //Normalize Formula
                        Normalize();
                        //Add the elements and their atomic weights to a dictionary
                        MatchCollection matches = Regex.Matches(Formula, @"([A-Z][a-z]*)");
                        string[] elementLetter = matches.Cast<Match>().Select(m => m.Value).ToArray();
                        foreach (DataRow row in ds.Tables["Elements"].Rows)
                        {
                            eleDict.Add(row.Field<String>("Symbol"), row.Field<Double>("AtomicWeight"));
                        }

                        //Group multiple element occurances in words array together
                        var groups = elementLetter.GroupBy(v => v);
                        foreach (var group in groups)
                        {

                            var numAlpha = new Regex("(?<Alpha>[A-Z][a-z]*)");
                            var match = numAlpha.Match(group.Key);
                            var alpha = match.Groups["Alpha"].Value;

                            //Check is the element dictionary contains the element from Formula
                            if (eleDict.ContainsKey(alpha))
                            {
                                //get the molecular weight
                                double molWeight = eleDict[alpha];
                                //get number of elements
                                int eleNumber = normDict[alpha];

                                //add to totals list
                                double total = molWeight * eleNumber;

                                calcTotals.Add(total);
                            }
                        }
                        //Sum all the totals in the calcTotals list and return newTotal
                        newTotal = calcTotals.Sum();

                        return newTotal;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error:" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 0.0;
                    }
                }
                else
                {
                    return 0.0;
                }
            }
        }

        //Normalize Formula
        public void Normalize()
        {
            //Check to see if the Formula is valid
            if (Valid() == true)
            {
                try
                {
                    normDict.Clear();
                    //Split up and count each word by capital letters
                    MatchCollection matches = Regex.Matches(Formula, @"([A-Z][a-z]*[0-9]*)");
                    string[] words = matches.Cast<Match>().Select(m => m.Value).ToArray();

                    // count how many times something occurs
                    var groups = words.GroupBy(v => v);
                    foreach (var group in groups)
                    {
                        //Go through the newNormalize again and split it up on letters and numbers
                        var numAlpha = new Regex("(?<Alpha>[A-Z][a-z]*)(?<Numeric>[0-9]*)");
                        var match = numAlpha.Match(group.Key);
                        var alpha = match.Groups["Alpha"].Value;
                        var num = match.Groups["Numeric"].Value;

                        //Check to see if the element has any numbers 
                        if (num == "")
                        {
                            //Check if the element exists in the dictionary 
                            if (normDict.ContainsKey(alpha))
                            {
                                //Get the value of the Key present in the dictionary and add it's count to it
                                int value = normDict[alpha];
                                normDict[alpha] = value + group.Count();
                            }
                            else
                                //Add the new element with their count to the dictionary
                                normDict.Add(alpha.ToString(), group.Count());
                        }
                        else
                        {
                            //Check if the element is in the dictionary
                            if (normDict.ContainsKey(alpha))
                            {
                                //get the element number that exists in the dictionary
                                int value = normDict[alpha];
                                //Add value + number and multiply it by group.Count()
                                normDict[alpha] = value + Convert.ToInt32(num) * group.Count();
                            }
                            else
                                //Otherwise add to the dictionary
                                normDict.Add(alpha.ToString(), Convert.ToInt32(num) * group.Count());
                        }

                    }

                    //Print the key back to the combobox
                    Formula = "";
                    foreach (KeyValuePair<string, int> sample in normDict)
                    {
                        if (sample.Value > 1)
                        {
                            Formula += sample.Key.ToString() + sample.Value.ToString();
                        }
                        else
                        {
                            Formula += sample.Key.ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Formula = Formula;
            }
        }

        public bool Save()
        {
            //Save Formula, and newTotal to the database
            try
            {
                if (Valid() == true)
                {
                    //Check to see if formula already exists in the database
                    //check if Formula is in formula array
                    Boolean contains = formList.Contains(Formula);
                    if (contains == true)
                    {
                        return false;
                    }
                    else
                    {
                        //Open up the database and insert Formula and newTotal into the database
                        string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Stoichiometry.accdb";
                        OleDbConnection conn = new OleDbConnection(connString);
                        conn.Open();
                        OleDbCommand cmd = conn.CreateCommand();
                        OleDbCommand cmdSelect = conn.CreateCommand();
                        cmd.CommandText = "INSERT INTO Molecules (Formula, MolecularWeight) VALUES ('" + Formula + "'," + newTotal + ")";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //add formula to the list
            return true;
        }

        public string[] FormulasList
        {
            get
            {
                try
                {
                    InitializeMolecule();
                    formList.Clear();
                    // Populate the Molecules list
                    foreach (DataRow row in ds.Tables["Molecules"].Rows)
                    {
                        formList.Add(row.Field<String>("Formula"));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return formList.ToArray();
            }
        }
        // C'tor
        public Molecule() { }
        public Molecule(string formula)
        {
            Formula = formula;
        }
        ~Molecule() { }

        //Open up the database and retrieve element intformation
        private DataSet InitializeElement()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Stoichiometry.accdb");
                odba = new OleDbDataAdapter();
                odba.SelectCommand = new OleDbCommand("SELECT * FROM Elements", conn);
                odba.InsertCommand = new OleDbCommand("INSERT INTO Elements (Formula) VALUES(@Formula)", conn);

                odba.InsertCommand.Parameters.Add("@Formula", OleDbType.VarChar, -1, "Formula");
                ds = new DataSet();
                conn.Open();
                odba.Fill(ds, "Elements");
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error no database connection!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }

        //Open the database and retrieve molecule information
        private DataSet InitializeMolecule()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Stoichiometry.accdb");
                odba = new OleDbDataAdapter();
                odba.SelectCommand = new OleDbCommand("SELECT Formula FROM Molecules", conn);
                odba.InsertCommand = new OleDbCommand("INSERT INTO Molecules (Formula) VALUES(@Formula)", conn);
                odba.InsertCommand.Parameters.Add("@Formula", OleDbType.VarChar, -1, "Formula");

                ds = new DataSet();
                conn.Open();
                odba.Fill(ds, "Molecules");
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error no database connection!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }

    }
}