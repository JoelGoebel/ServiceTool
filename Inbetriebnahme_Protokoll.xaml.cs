using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;


namespace ServiceTool
{
    /// <summary>
    /// Interaktionslogik für UserControl1.xaml
    /// </summary>
    public partial class Inbetriebnahme_Protokoll : UserControl
    {
        public Inbetriebnahme_Protokoll(bool isFirstLoad = false)
        {
            InitializeComponent();
            FillSiteSwitchComboBox();//add all The Sites to the Combobox            
        }

        public bool FirstSiteLoadFinished = GlobalVariables._FirstSiteLoadFinished ;
        
        public void FillSiteSwitchComboBox()
        {
            //Deactivate the SelectionChanged Event to prevent the event from being triggered when the ComboBox is filled
            CB_SeiteAuswählen.SelectionChanged -= SiteSelected;
            bool StartSiteSelected = GlobalVariables.StartSiteSelected;
            //If a Maschine is selected that is not MRS or Jump, add it to the Combobox
            if (GlobalVariables.Maschiene_1 != "" && GlobalVariables.Maschiene_1 != "MRS" && GlobalVariables.Maschiene_1 != "Jump")
            {
                CB_Item1_SeiteAuswählen.Content = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                CB_Item1_SeiteAuswählen.Visibility = Visibility.Visible;
                if(StartSiteSelected == false)
                {//If no Site is selected yet, select the first one
                    CB_SeiteAuswählen.SelectedIndex = 0;
                    StartSiteSelected = true;
                }
            }
            if (GlobalVariables.Maschiene_2 != "" && GlobalVariables.Maschiene_2 != "MRS" && GlobalVariables.Maschiene_2 != "Jump")
            {
                CB_Item2_SeiteAuswählen.Content = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                CB_Item2_SeiteAuswählen.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    CB_SeiteAuswählen.SelectedIndex = 1;
                    StartSiteSelected = true;
                }                
            }
            if (GlobalVariables.Maschiene_3 != "" && GlobalVariables.Maschiene_3 != "MRS" && GlobalVariables.Maschiene_3 != "Jump")
            {
                CB_Item3_SeiteAuswählen.Content = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                CB_Item3_SeiteAuswählen.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    CB_SeiteAuswählen.SelectedIndex = 2;
                    StartSiteSelected = true;
                }
            }
            if (GlobalVariables.Maschiene_4 != "" && GlobalVariables.Maschiene_4 != "MRS" && GlobalVariables.Maschiene_4 != "Jump")
            {
                CB_Item4_SeiteAuswählen.Content = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                CB_Item4_SeiteAuswählen.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    CB_SeiteAuswählen.SelectedIndex = 3;
                    StartSiteSelected = true;
                }
            }
            // Trandfer the Information to the GlobalVariables Class
            GlobalVariables.StartSiteSelected = StartSiteSelected;
            if(GlobalVariables._FirstSiteLoadFinished == false) 
            {
                // With the Dispatcher we can ensure that the FirstSiteLoadFinished is set to true after the UI is loaded
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    GlobalVariables._FirstSiteLoadFinished = true;
                    FirstSiteLoadFinished = GlobalVariables._FirstSiteLoadFinished;
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            }
            CB_SeiteAuswählen.SelectionChanged += SiteSelected; // Re-activate the SelectionChanged Event after the ComboBox is filled
        }


        public void SiteSelected(object sender, SelectionChangedEventArgs e)
        {
            //Check if the FirstSiteLoadFinished is false, if so, do not execute the code
            if (FirstSiteLoadFinished == false)
            {
                return;
            }
            //Deactivate the SelectionChanged Event to prevent the event from being triggered when the ComboBox is filled
            CB_SeiteAuswählen.SelectionChanged -= SiteSelected;

            MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
            string LastSelectedItem = CB_SeiteAuswählen.SelectionBoxItem.ToString(); // Safe the last selected item before changing it
            string selectedItem = CB_SeiteAuswählen.SelectedItem.ToString(); // Get the currently selected item from the ComboBox
            string selectedItemText = selectedItem.Substring(selectedItem.IndexOf(" ")+1);
            string ExcelFilePathLoad = "";
            string ExcelFilePathSave = "";

            //Set ExcelfilePath for saving last Selected Site
            if (LastSelectedItem == "" || LastSelectedItem == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "IbnP");
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2 )
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_2.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "IbnP");
             }
            else if (LastSelectedItem == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "IbnP");
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "IbnP");
            }
            //Set ExcelfilePath for loading new Selected Site
            if (selectedItemText == "" || selectedItemText == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP");
                tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_1;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_2.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP");
                tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_2;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP");
                tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_3;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP");
                tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_4;
            }
            //Acivate the SelectionChanged Event again
            CB_SeiteAuswählen.SelectionChanged += SiteSelected;       
        }



        public void Bearbeitung_deaktivieren()
        {
            string basisOrdner = AppDomain.CurrentDomain.BaseDirectory;
            string relativerPfad = @"Infos/TBXCB_mit_Zellen.xlsx";
            string filePath = System.IO.Path.Combine(basisOrdner, relativerPfad);

            // Initialisiere das zweidimensionale Array (wir wissen vorher nicht, wie viele Zeilen es gibt)
            string[,] Zellen_Textboxen;
            string[,] Zellen_Checkboxen;

            // Lade die Excel-Datei
            FileInfo fileInfo = new FileInfo(filePath);

            // Verwende EPPlus, um die Excel-Datei zu lesen
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Nimm das erste Arbeitsblatt
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Bestimme die Anzahl der gefüllten Zeilen
                int rowCount = worksheet.Dimension.Rows;

                // Initialisiere das zweidimensionale Array basierend auf der Anzahl der Zeilen
                Zellen_Textboxen = new string[rowCount, 2]; // 2 Spalten (Zelle, TextBox-Name)

                Zellen_Checkboxen = new string[20, 2];

                // Lese die Daten aus der ersten und zweiten Spalte
                for (int row = 2; row <= rowCount; row++)
                {
                    Zellen_Textboxen[row - 1, 0] = worksheet.Cells[row, 1].Text; // Spalte 1 (Zelle)
                    Zellen_Textboxen[row - 1, 1] = worksheet.Cells[row, 2].Text; // Spalte 2 (TextBox-Name)

                    // Den Namen der ersten TextBox aus dem Array auslesen
                    string TextBoxName = Zellen_Textboxen[row - 1, 1]; // Die zweite Spalte enthält die TextBox-Namen

                    // Verwende FindName, um das TextBox-Element im UserControl zu finden
                    TextBox TextBox_ = this.FindName(TextBoxName) as TextBox;

                    TextBox_.IsEnabled = false;

                    if (worksheet.Cells[row, 4].Text != "")
                    {
                        Zellen_Checkboxen[row - 1, 0] = worksheet.Cells[row, 4].Text;
                        Zellen_Checkboxen[row - 1, 1] = worksheet.Cells[row, 5].Text;

                        string CheckboxName = Zellen_Checkboxen[row - 1, 1];

                        CheckBox Checkbox_ = this.FindName(CheckboxName) as CheckBox;

                        Checkbox_.IsEnabled = false;
                    }

                }

            }
        }


       

    }
}
