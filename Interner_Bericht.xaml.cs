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

namespace ServiceTool
{
    /// <summary>
    /// Interaktionslogik für Window3.xaml
    /// </summary>
    public partial class Interner_Bericht : UserControl
    {
        private bool _isInitialized = false;
        private bool StartSiteselected = false;
        public Interner_Bericht()
        {
            InitializeComponent();
            SetNameOfComboBoxItem();
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _isInitialized = true;
                if(CB_Einheit_M.SelectedItem == null)
                {
                    CB_Einheit_M.Text = "Nm";
                }
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);            
        }
        /// Funktion um die Maschinennamen der ComboBox zuzuweisen
        private void SetNameOfComboBoxItem()
        {
            if (GlobalVariables.Maschiene_1 != "")
            {
                cbItem_InBe_1.Content = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                cbItem_InBe_1.Visibility = Visibility.Visible;
                if(StartSiteselected == false)
                {
                    cb_SeitenWechselInternerBericht.SelectedItem = cbItem_InBe_1;
                    StartSiteselected = true;
                }
            }
            if(GlobalVariables.Maschiene_2 != "")
            {
                cbItem_InBe_2.Content = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                cbItem_InBe_2.Visibility = Visibility.Visible;
                if (StartSiteselected == false)
                {
                    cb_SeitenWechselInternerBericht.SelectedItem = cbItem_InBe_2;
                    StartSiteselected = true;
                }
            }
            if (GlobalVariables.Maschiene_3 != "")
            {
                cbItem_InBe_3.Content = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                cbItem_InBe_3.Visibility = Visibility.Visible;
                if (StartSiteselected == false)
                {
                    cb_SeitenWechselInternerBericht.SelectedItem = cbItem_InBe_3;
                    StartSiteselected = true;
                }
            }
            if (GlobalVariables.Maschiene_4 != "")
            {
                cbItem_InBe_4.Content = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                cbItem_InBe_4.Visibility = Visibility.Visible;
                if (StartSiteselected == false)
                {
                    cb_SeitenWechselInternerBericht.SelectedItem = cbItem_InBe_4;
                    StartSiteselected = true;
                }
            }   
        }

        private void SiteSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized == false)
            {
                return;
            }

            cb_SeitenWechselInternerBericht.SelectionChanged -= SiteSelectionChanged;

            MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
            string LastSelectedItem = cb_SeitenWechselInternerBericht.SelectionBoxItem.ToString();
            string selectedItem = cb_SeitenWechselInternerBericht.SelectedItem.ToString();
            string selectedItemText = selectedItem.Substring(selectedItem.IndexOf(" ") + 1);
            string ExcelFilePathLoad = "";
            string ExcelFilePathSave = "";

            
            //Set ExcelfilePath for saving last Selected Site
            if (LastSelectedItem == "" || LastSelectedItem == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "Interner_Bericht");
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_2.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "Interner_Bericht");
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_3.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "Interner_Bericht");
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_4.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "Interner_Bericht");
            }
            //Set ExcelfilePath for loading new Selected Site
            if (selectedItemText == "" || selectedItemText == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "Interner_Bericht");
            }
            else if (selectedItemText == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_2.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "Interner_Bericht");
            }
            else if (selectedItemText == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_3.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "Interner_Bericht");
            }
            else if (selectedItemText == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_4.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "Interner_Bericht");
            }

            cb_SeitenWechselInternerBericht.SelectionChanged += SiteSelectionChanged;
        }


        private void CB_Einheit_M_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string result = "";
            if (_isInitialized)
            {
                ComboBox cb = sender as ComboBox;
                if (cb.SelectedItem == null) 
                {
                    result = "Nm";
                    CB_Einheit_M.SelectionChanged -= CB_Einheit_M_SelectionChanged;
                    CB_Einheit_M.Text = result;
                    CB_Einheit_M.SelectionChanged += CB_Einheit_M_SelectionChanged;
                }
                else
                {
                    string input = cb.SelectedItem.ToString();
                    int index = input.IndexOf(' ');
                    result = index >= 0 ? input.Substring(index + 1) : input;
                }
                    
                
                CB_Einheit_T1.Text = result; 
                CB_Einheit_B1.Text = result;
                CB_Einheit_T2.Text = result;
                CB_Einheit_T3.Text = result;
                CB_Einheit_B2.Text = result;
                CB_Einheit_B3.Text = result;
                CB_Einheit_T4.Text = result;
                CB_Einheit_B4.Text = result;
                CB_Einheit_TB0.Text = result;
                CB_Einheit_D0.Text = result;
                CB_Einheit_Sonstige.Text = result;
                
            }
        }

        private void CB_Einheit_M_SelectionChanged1(object sender, SelectionChangedEventArgs e)
        {
            throw new NotImplementedException();
        }
    }
}
