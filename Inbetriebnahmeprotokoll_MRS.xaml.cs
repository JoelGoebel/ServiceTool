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
    /// Interaktionslogik für Inbetriebnahmeprotokoll_MRS.xaml
    /// </summary>
    public partial class Inbetriebnahmeprotokoll_MRS : UserControl
    {
        private bool StartSiteSelected = false;
        private bool _isInitialized = false;
        public Inbetriebnahmeprotokoll_MRS()
        {
            InitializeComponent();
            FillSiteSwitchComboBox();

            Dispatcher.BeginInvoke(new Action(() =>
            {
                _isInitialized = true;
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            GlobalVariables.LastSelectedItem_MRS = cb_SiteSwitchIbnP_MRS.SelectionBoxItem.ToString();
        }

        public void FillSiteSwitchComboBox()
        {

            cb_SiteSwitchIbnP_MRS.SelectionChanged -= SiteSelectionChanged;
            //Wenn in den ServiceAnforderungen eine Maschine eingetragen wurde wird der Typ in den Klasse GlobalVariables gespeichert
            if (GlobalVariables.Maschiene_1 != "" && (GlobalVariables.Maschiene_1 == "MRS" || GlobalVariables.Maschiene_1 == "Jump"))
            {
                cbItem_Site1.Content = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                cbItem_Site1.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    cb_SiteSwitchIbnP_MRS.SelectedItem = cbItem_Site1;
                    StartSiteSelected = true;
                }
            }
            if (GlobalVariables.Maschiene_2 != "" && (GlobalVariables.Maschiene_2 == "MRS" || GlobalVariables.Maschiene_2 == "Jump"))
            {
                cbItem_Site2.Content = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                cbItem_Site2.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    cb_SiteSwitchIbnP_MRS.SelectedItem = cbItem_Site2; ;
                    StartSiteSelected = true;
                }
            }
            if (GlobalVariables.Maschiene_3 != "" && (GlobalVariables.Maschiene_3 == "MRS" || GlobalVariables.Maschiene_3 == "Jump"))
            {
                cbItem_Site3.Content = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                cbItem_Site3.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    cb_SiteSwitchIbnP_MRS.SelectedItem = cbItem_Site3;
                    StartSiteSelected = true;
                }
            }
            if (GlobalVariables.Maschiene_4 != "" && (GlobalVariables.Maschiene_4 == "MRS" || GlobalVariables.Maschiene_4 == "Jump"))
            {
                cbItem_Site4.Content = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                cbItem_Site4.Visibility = Visibility.Visible;
                if (StartSiteSelected == false)
                {
                    cb_SiteSwitchIbnP_MRS.SelectedItem = cbItem_Site4;
                    StartSiteSelected = true;
                }
            }
            cb_SiteSwitchIbnP_MRS.SelectionChanged += SiteSelectionChanged;
        }

        private void SiteSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized == false)
            {
                return;
            }

            cb_SiteSwitchIbnP_MRS.SelectionChanged -= SiteSelectionChanged;

            MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
            string LastSelectedItem = cb_SiteSwitchIbnP_MRS.SelectionBoxItem.ToString();
            GlobalVariables.LastSelectedItem_MRS = LastSelectedItem;
            string selectedItem = cb_SiteSwitchIbnP_MRS.SelectedItem.ToString();
            string selectedItemText = selectedItem.Substring(selectedItem.IndexOf(" ") + 1);
            string ExcelFilePathLoad = "";
            string ExcelFilePathSave = "";
            string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS.png");
            string ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS.png");


            //Set ExcelfilePath for saving last Selected Site
            if (LastSelectedItem == "" || LastSelectedItem == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "IbnP_MRS");
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_2.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_2.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_2.png");
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_3.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_3.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_3.png");
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_4.xlsx");
                mainWindow.speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_4.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_4.png");
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                mainWindow.SaveSignatureAsImage(ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            //Set ExcelfilePath for loading new Selected Site
            if (selectedItemText == "" || selectedItemText == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP_MRS");
                tb_ExtruderTyp_ibnProtokoll_MRS.Text = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                tb_Seriennummer_ibnProtokoll_MRS.Text = GlobalVariables.MaschinenNr_1;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_2.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP_MRS");
                tb_ExtruderTyp_ibnProtokoll_MRS.Text = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                tb_Seriennummer_ibnProtokoll_MRS.Text = GlobalVariables.MaschinenNr_2;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_3.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP_MRS");
                tb_ExtruderTyp_ibnProtokoll_MRS.Text = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                tb_Seriennummer_ibnProtokoll_MRS.Text = GlobalVariables.MaschinenNr_3;
            }
            else if (selectedItemText == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_4.xlsx");
                mainWindow.Laden(ExcelFilePathLoad, "IbnP_MRS");
                tb_ExtruderTyp_ibnProtokoll_MRS.Text = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                tb_Seriennummer_ibnProtokoll_MRS.Text = GlobalVariables.MaschinenNr_4;
            }

            tb_Kunde_ibnProtokoll_MRS.Text = GlobalVariables.Kunde;

            cb_SiteSwitchIbnP_MRS.SelectionChanged += SiteSelectionChanged;
        }
    }
}
