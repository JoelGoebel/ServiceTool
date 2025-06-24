using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
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
using FuzzySharp;

namespace ServiceTool
{
    /// <summary>
    /// Interaktionslogik für UserControl1.xaml
    /// </summary>
    public partial class Startseite : UserControl
    {
        private MainWindow mw;
        public Startseite()//MainWindow mainWindow
        {
            
            InitializeComponent();
            mw = (MainWindow)Application.Current.MainWindow;//get the MainWindow Instance in order to acces its Functions
            //Insert the DataTable into the DataGrid
            if (GlobalVariables.dt != null)
            {
                var filteredData = GlobalVariables.dt.AsEnumerable();
                dataGrid.ItemsSource = filteredData.CopyToDataTable().DefaultView;
            }
        }

        private void btn_AuftragsNr_suchen_Click(object sender, RoutedEventArgs e)
        {
            mw._isInitialized = false;
            ClearGlobalVariables();//Reset all Information in GlobalVariables

            //Check if User Enterd Auftragsnummer
            if (tb_Auftragsnummer.Text == "" )
            {
                MessageBox.Show("Bitte geben Sie eine Auftragsnummer ein.");
                return;
            }
            //Check if Auftragsnummer is Only Numbers
            if (!int.TryParse(tb_Auftragsnummer.Text, out int zahl))
            {
                MessageBox.Show("Die Auftragsnummer darf nur aus Zahlen bestehen");
                tb_Auftragsnummer.Text = "";
                return;
            }

            // Safe the Auftragsnummer in GlobalVariables for later use
            GlobalVariables.AuftragsNR = tb_Auftragsnummer.Text;
            string Auftragsnummer = GlobalVariables.AuftragsNR;

            PfadeFestlegen(); //Set all Paths for the current OrderNo

            if (GlobalVariables.dt != null) 
            { 
                Datenabgleich_Database();
            }
            AuftragsOrdnerErstellen();

            CopyVorlagenInAuftragsOrdner();

            DokumenteEinblenden();

            if (GlobalVariables.Online_or_Offline) 
            {
                OfflineDatenAbgleich();
            }

            mw.CheckAllRadioButtons();

            if(GlobalVariables.Sprache_Kunde != "D")
            {
                mw.CB_Sprache_auswahl.Text = "Englisch";
            }
            else
            {
                mw.CB_Sprache_auswahl.Text = "Deutsch";
            }

            mw.tb_KundenEmail.Text = GlobalVariables.CustomerEmail;

            mw._isInitialized = true;
        }

        //TO OfflineDatenAbgleich WE
        //DO OfflineDatenArchivieren
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void OfflineDatenAbgleich()
        {
            //Set Local Paths to work offline
            string Pfad_DokumentOrdner = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string Pfad_Servicetool_Lokal = string.Format(Properties.Resources.Pfad_AuftragsOrdner_Off, GlobalVariables.AuftragsNR);
            string Temp_Pfad_AuftragsOrdner = System.IO.Path.Combine(Pfad_DokumentOrdner, Pfad_Servicetool_Lokal);

            if (Directory.Exists(Temp_Pfad_AuftragsOrdner))//Check if the Local Order Folder exists
            {
                foreach(string datei in Directory.GetFiles(Temp_Pfad_AuftragsOrdner, "*.xls*"))// Do for each Excel Data in the Local Order Folder
                {
                    string dateiName = System.IO.Path.GetFileName(datei); //get the Filename of the Excel Data
                    string quellDateiPfad = System.IO.Path.Combine(Temp_Pfad_AuftragsOrdner, dateiName);
                    string zielDateiPfad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, dateiName);
                    DateTime LastChange_OfflineData = File.GetLastWriteTime(quellDateiPfad);//check the last change Date of the Local Data
                    DateTime LastChange_OnlineData = File.GetLastWriteTime(zielDateiPfad);// check the last change Date of the Online Data

                    if (File.Exists(zielDateiPfad) && dateiName != "Service_Anforderungen.xlsx")
                    { // if the Online Data exists and is not the Service_Anforderungen.xlsx
                        if (LastChange_OfflineData>=LastChange_OnlineData)
                        {//If the Local Data is newer than the Online Data
                            File.Copy(quellDateiPfad,zielDateiPfad,true);
                            OfflineDatenArchivieren(quellDateiPfad, dateiName);
                        }
                    }
                    else
                    {
                        File.Copy(quellDateiPfad, zielDateiPfad,true);
                    }

                }
            }

        }
        // Unterfunktion OfflineDatenArchivieren
        private void OfflineDatenArchivieren(string quellDateiPfad, string DateiName)
        {
            string NutzerName = Environment.UserName;
            string Nutzer_ArchivPfad = System.IO.Path.Combine(Properties.Resources.ArchievePfad,NutzerName);
            if (!Directory.Exists(Nutzer_ArchivPfad)) { Directory.CreateDirectory(Nutzer_ArchivPfad); }
            string[] dateien = Directory.GetFiles(Nutzer_ArchivPfad, "*.xls*");
            int Anzahl_ArchivierteDaten = dateien.Length;

            if (Anzahl_ArchivierteDaten < 10)
            {//If less than 10 files are archived, copy the file to the archive
                string NutzerArchiveDateiPfad = System.IO.Path.Combine(Nutzer_ArchivPfad, DateiName);
                if(File.Exists(NutzerArchiveDateiPfad)) { File.Delete(NutzerArchiveDateiPfad); }
                File.Copy(quellDateiPfad, NutzerArchiveDateiPfad);
                if (File.Exists(quellDateiPfad)) { File.Delete(quellDateiPfad); }
            }
            else
            {//If more than 10 files are archived, delete the oldest file and copy the new file to the archive
                DirectoryInfo VerzeichnisInfos = new DirectoryInfo(Nutzer_ArchivPfad);
                FileInfo aeltesteDatei = VerzeichnisInfos.GetFiles().OrderBy(datei => datei.LastWriteTime).FirstOrDefault();
                File.Delete(aeltesteDatei.FullName);
                File.Move(quellDateiPfad, Nutzer_ArchivPfad);
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void DokumenteEinblenden()
        {// Set the Visibility of the RadioButtons and Buttons in the MainWindow to Visible.
            mw.rbt_InbetriebnahmeProtokoll.Visibility = Visibility.Visible;
            mw.rbt_InbetriebnahmeProtokoll_MRS.Visibility = Visibility.Visible;            
            mw.rbt_ServiceAnforderung.Visibility = Visibility.Visible;
            mw.rbt_Stundennachweis.Visibility = Visibility.Visible;
            mw.rbt_InternerBericht.Visibility = Visibility.Visible;
            mw.rbt_Speichern.Visibility = Visibility.Visible;
            mw.btn_Anhaege_hinzufuegen.Visibility = Visibility.Visible;
            mw.btn_Auftrag_Download.Visibility = Visibility.Visible;
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void CopyVorlagenInAuftragsOrdner()
        {
            // Copy all template files from the source folder to the order folder
            foreach (string datei in Directory.GetFiles(GlobalVariables.Pfad_QuelleVorlagen))
            {
                string dateiName = System.IO.Path.GetFileName(datei);
                string zielDateiPfad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, dateiName);

                // Check if Data already exists in the target folder
                if (!File.Exists(zielDateiPfad))
                {
                    File.Copy(datei, zielDateiPfad);
                    Console.WriteLine($"Die Datei '{dateiName}' wurde kopiert.");
                }
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void AuftragsOrdnerErstellen()
        {
            //Generate the Folder for the current OrderNo if it does not exist
            if (!Directory.Exists(GlobalVariables.Pfad_AuftragsOrdner))
            {
                Directory.CreateDirectory(GlobalVariables.Pfad_AuftragsOrdner);
                Directory.CreateDirectory(GlobalVariables.Pfad_Anhaenge);
                Directory.CreateDirectory(GlobalVariables.Pfad_Unterschriften);
                Directory.CreateDirectory(GlobalVariables.Pfad_Anhaenge + @"\Fotos");
                Console.WriteLine($"Der Ordner '{GlobalVariables.Pfad_AuftragsOrdner}' wurde erstellt.");
            }
            else
            {
                Console.WriteLine($"Der Ordner '{GlobalVariables.Pfad_AuftragsOrdner}' existiert bereits.");
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void Datenabgleich_Database()
        {
            DataTable DB_Serviceaufträge = GlobalVariables.dt;
            int row_count = DB_Serviceaufträge.Rows.Count;// Get the number of rows in the DataTable

            for (int i = 0; i < row_count; i++)
            {
                string DB_AuftragsNR = Convert.ToString(DB_Serviceaufträge.Rows[i][1]);
                //Search for the Order Number in the DataTable if found set the GlobalVariables
                if (DB_AuftragsNR == GlobalVariables.AuftragsNR)
                {
                    GlobalVariables.Kunde = Convert.ToString(DB_Serviceaufträge.Rows[i][5]);
                    GlobalVariables.KundenNummer = Convert.ToString(DB_Serviceaufträge.Rows[i][4]);
                    GlobalVariables.Sprache_Kunde = Convert.ToString(DB_Serviceaufträge.Rows[i][3]);
                    GlobalVariables.Land = Convert.ToString(DB_Serviceaufträge.Rows[i][6]);
                    GlobalVariables.Anschrift_1 = Convert.ToString(DB_Serviceaufträge.Rows[i][10]);
                    string temp_PLZ = Convert.ToString(DB_Serviceaufträge.Rows[i][8]);
                    string temp_Ort = Convert.ToString(DB_Serviceaufträge.Rows[i][9]);
                    GlobalVariables.Anschrift_2 = $@"{temp_PLZ} {temp_Ort}";
                    GlobalVariables.auftraginDB = true;
                    GlobalVariables.CustomerEmail = Convert.ToString(DB_Serviceaufträge.Rows[i][14]);
                }
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void PfadeFestlegen()
        {
            string basisOrdner = AppDomain.CurrentDomain.BaseDirectory;
            string relativerPfad = @"Vorlagen";
            GlobalVariables.Pfad_QuelleVorlagen = System.IO.Path.Combine(basisOrdner, relativerPfad);

            //Check if the Server is reachable
            if (GlobalVariables.Online_or_Offline)
            {
                // Set Online Paths
                GlobalVariables.Pfad_AuftragsOrdner = string.Format(Properties.Resources.Pfad_AuftragsOrdner_On, GlobalVariables.AuftragsNR);
                GlobalVariables.Pfad_Anhaenge = string.Format(Properties.Resources.Pfad_Anhaenge_On, GlobalVariables.AuftragsNR);
                GlobalVariables.Pfad_Unterschriften = string.Format(Properties.Resources.Pfad_Signatures_On, GlobalVariables.AuftragsNR);
            }
            else
            {
                //Set Offline Paths
                string Pfad_DokumentOrdner= Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string Pfad_Servicetool_Lokal = string.Format(Properties.Resources.Pfad_AuftragsOrdner_Off, GlobalVariables.AuftragsNR);
                GlobalVariables.Pfad_AuftragsOrdner = System.IO.Path.Combine(Pfad_DokumentOrdner, Pfad_Servicetool_Lokal);
                GlobalVariables.Pfad_Anhaenge = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge");
                GlobalVariables.Pfad_Unterschriften = System.IO.Path.Combine(GlobalVariables.Pfad_Anhaenge, "Unterschriften");

            }
        }

        private void ResetSearch(object sender, RoutedEventArgs e)
        {
            tb_Land_Startseite.Text = "";
            tb_KundenFirmenName_Startseite.Text = "";
            dp_AbDatum_Startseite.Text = "";
            var filteredData = GlobalVariables.dt.AsEnumerable();
            dataGrid.ItemsSource = filteredData.CopyToDataTable().DefaultView;
        }

        private void OnSearchClick(object sender, RoutedEventArgs e)
        {
            //Check if at least one search criteria is entered else Funktion
            if (tb_Land_Startseite.Text == "" && tb_KundenFirmenName_Startseite.Text == "" && dp_AbDatum_Startseite.Text == "")
            {
                MessageBox.Show("Bitte geben Sie mindestens ein Suchkriterium ein.");
                return;
            }
            //Place the DataView of the DataGrid into a variable
            var view = dataGrid.ItemsSource as DataView;
            DataTable table = view?.ToTable(); // Contains all sorted Data

            var filteredData = table?.AsEnumerable(); // als IEnumerable<DataRow>

            if (!string.IsNullOrEmpty(tb_Land_Startseite.Text))
            {// if Land is enterd as a search criteria and look up the DataTable for it
                filteredData = filteredData.Where(row => row.Field<string>("Land").ToUpper().Contains(tb_Land_Startseite.Text.ToUpper()));
            }

            if (!string.IsNullOrEmpty(tb_KundenFirmenName_Startseite.Text))
            {// if KundenFirmenName is entered as a search criteria and look up the DataTable for it
                filteredData = filteredData.Where(row => row.Field<string>("A0Name1").ToUpper().Contains(tb_KundenFirmenName_Startseite.Text.ToUpper()));
            }

            if (dp_AbDatum_Startseite.SelectedDate.HasValue)
            {//Check if a date is selected
                if(dp_AbDatum_Startseite.SelectedDate.Value <= DateTime.Today)
                { // Check if the selected date is not in the future. if not filter the DataTable for it
                    var abDatum = dp_AbDatum_Startseite.SelectedDate.Value;
                    filteredData = filteredData.Where(row => row.Field<DateTime>("Belegdatum") >= abDatum);
                }
                else
                {// If future date is selected show a message box and return
                    MessageBox.Show("Das Datum darf nicht in der Zukunft liegen.");
                    return;
                }
            }

            if (filteredData == null || !filteredData.Any())
            {//if no data is found show a message box and return
                dataGrid.BorderBrush = Brushes.Red;
                dataGrid.BorderThickness = new Thickness(4);
                MessageBox.Show("Keine Daten gefunden.");
                dataGrid.BorderBrush = Brushes.Black;
                dataGrid.BorderThickness = new Thickness(1);
                return;
            }

            try
            {//Try to insert the filtered data into the DataGrid
                dataGrid.ItemsSource = filteredData.CopyToDataTable().DefaultView;
            }
            catch (InvalidOperationException)
            {//if it doesn,t work it means that no data is found
                dataGrid.BorderBrush = Brushes.Red;
                dataGrid.BorderThickness = new Thickness(4);
                MessageBox.Show("Keine Daten gefunden");
                dataGrid.BorderBrush = Brushes.Black;
                dataGrid.BorderThickness = new Thickness(1);
                throw;
            }
           
        }
        private void InsertAuftragsNrWhenRowSelected(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.SelectedItem == null)
            {
                return;
            }
            DataRowView row = (DataRowView)dataGrid.SelectedItem;
            tb_Auftragsnummer.Text = row.Row.ItemArray[1].ToString();
            // Set the Cursor to the TextBox
            tb_Auftragsnummer.Focus();

            // Select all the Text in the TextBox
            tb_Auftragsnummer.SelectAll();
        }

        private void AuftragsNrSucheStarten(object sender, KeyEventArgs e)
        {//Event for pressing Enter in the Auftragsnummer TextBox
            if (e.Key == Key.Enter)
            {
                btn_AuftragsNr_suchen_Click(sender, e);
            }
        }

        private void StarteTabellenFilter(object sender, KeyEventArgs e)
        {//Event for pressing Enter in the Search TextBoxes
            if (e.Key == Key.Enter)
            {
                OnSearchClick(sender, e);
            }
        }
    
        private void ClearGlobalVariables()
        {//Reset all GlobalVariables to their default values
            GlobalVariables.Kunde = "";
            GlobalVariables.KundenNummer = "";
            GlobalVariables.Sprache_Kunde = "";
            GlobalVariables.auftraginDB = false;
            GlobalVariables.ServiceTechnicker = "";
            GlobalVariables.Ansprechpartner = "";
            GlobalVariables.Anreise = "";
            GlobalVariables.Land = "";
            GlobalVariables.Anschrift_1 = "";
            GlobalVariables.Anschrift_2 = "";
            GlobalVariables.Material = "";
            GlobalVariables.Maschiene_1 = "";
            GlobalVariables.Maschiene_2 = "";
            GlobalVariables.Maschiene_3 = "";
            GlobalVariables.Maschiene_4 = "";
            GlobalVariables.Baugroeße_1 = "";
            GlobalVariables.Baugroeße_2 = "";
            GlobalVariables.Baugroeße_3 = "";
            GlobalVariables.Baugroeße_4 = "";
            GlobalVariables.MaschinenNr_1 = "";
            GlobalVariables.MaschinenNr_2 = "";
            GlobalVariables.MaschinenNr_3 = "";
            GlobalVariables.MaschinenNr_4 = "";
            GlobalVariables.Signatur_IBN_1 = false;
            GlobalVariables.Signatur_IBN_2 = false;
            GlobalVariables.Signatur_IBN_3 = false;
            GlobalVariables.Signatur_IBN_4 = false;
        }
    }
}
