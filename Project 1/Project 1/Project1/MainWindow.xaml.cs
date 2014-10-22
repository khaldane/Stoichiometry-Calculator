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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Stoichiometry;

namespace Project1
{

    public partial class MainWindow : Window
    {
        //Global Scope
        Molecule mole = new Molecule();

        public MainWindow()
        {

            try
            {
                InitializeComponent();
                //Populate the formula combobox with all the user saved molecules from the database
                foreach (string molecule in mole.FormulasList)
                {
                    FormulaComboBox.Items.Add(molecule);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error no database connection!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }


        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void FormulaComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        //When the normalize button is clicked it will call the normalize function and populate the combobox with returned value
        private void Normalize_Click(object sender, RoutedEventArgs e)
        {
            mole.Formula = FormulaComboBox.Text;
            string prevFormula = FormulaComboBox.Text;
            mole.Normalize();
            FormulaComboBox.Text = mole.Formula;
            if (mole.Valid() == false)
            {
                MessageBox.Show("Invalid Characters in Formula, please try again!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
                MessageBox.Show("Normalization Complete!", "Complete", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        //Calculates the Formula from the textbox
        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            //Set Formula from string in combobox
            mole.Formula = FormulaComboBox.Text;
            bool test = mole.Valid();
            //Set the WeightTotal total to the MolecularWeight calculated
            WeightTotal.Text = mole.MolecularWeight.ToString();
            //If Molecular weight does not equal zero save to database otherwise do nothing
            if (mole.MolecularWeight.ToString() != "0")
            {
                MessageBoxResult dr = MessageBox.Show("Would you like to permanently store this formula in the database?", "Would you like to save?", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (dr == MessageBoxResult.Yes)
                {
                    mole.Formula = FormulaComboBox.Text;
                    //Check to see if the save was completed
                    if (mole.Save() == true)
                    {
                        //Refresh the combobox FormulaList
                        FormulaComboBox.Items.Clear();
                        foreach (string molecule in mole.FormulasList)
                        {
                            FormulaComboBox.Items.Add(molecule);
                        }
                        MessageBox.Show("Element has been saved to the database!", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Element already exists in the database!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        //Close window
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
