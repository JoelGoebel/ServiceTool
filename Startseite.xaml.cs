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
            mw = (MainWindow)Application.Current.MainWindow;
            var filteredData = GlobalVariables.dt.AsEnumerable();
            dataGrid.ItemsSource = filteredData.CopyToDataTable().DefaultView;

        }

        private void btn_AuftragsNr_suchen_Click(object sender, RoutedEventArgs e)
        {
            mw._isInitialized = false;
            ClearGlobalVariables();

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
            

            GlobalVariables.AuftragsNR = tb_Auftragsnummer.Text;
            string Auftragsnummer = GlobalVariables.AuftragsNR;

            PfadeFestlegen();

            Datenabgleich_Database();

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
            mw._isInitialized = true;
        }

        //TO OfflineDatenAbgleich WE
        //DO OfflineDatenArchivieren
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void OfflineDatenAbgleich()
        {
            string Pfad_DokumentOrdner = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string Pfad_Servicetool_Lokal = string.Format(Properties.Resources.Pfad_AuftragsOrdner_Off, GlobalVariables.AuftragsNR);
            string Temp_Pfad_AuftragsOrdner = System.IO.Path.Combine(Pfad_DokumentOrdner, Pfad_Servicetool_Lokal);

            if (Directory.Exists(Temp_Pfad_AuftragsOrdner))
            {
                foreach(string datei in Directory.GetFiles(Temp_Pfad_AuftragsOrdner, "*.xls*"))
                {
                    string dateiName = System.IO.Path.GetFileName(datei);
                    string quellDateiPfad = System.IO.Path.Combine(Temp_Pfad_AuftragsOrdner, dateiName);
                    string zielDateiPfad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, dateiName);
                    DateTime LastChange_OfflineData = File.GetLastWriteTime(quellDateiPfad);
                    DateTime LastChange_OnlineData = File.GetLastWriteTime(zielDateiPfad);

                    if (File.Exists(zielDateiPfad) && dateiName != "Service_Anforderungen.xlsx") 
                    { 
                        if (LastChange_OfflineData>=LastChange_OnlineData)
                        {
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
            {
                string NutzerArchiveDateiPfad = System.IO.Path.Combine(Nutzer_ArchivPfad, DateiName);
                if(File.Exists(NutzerArchiveDateiPfad)) { File.Delete(NutzerArchiveDateiPfad); }
                File.Copy(quellDateiPfad, NutzerArchiveDateiPfad);
                if (File.Exists(quellDateiPfad)) { File.Delete(quellDateiPfad); }
            }
            else
            {
                DirectoryInfo VerzeichnisInfos = new DirectoryInfo(Nutzer_ArchivPfad);
                FileInfo aeltesteDatei = VerzeichnisInfos.GetFiles().OrderBy(datei => datei.LastWriteTime).FirstOrDefault();
                File.Delete(aeltesteDatei.FullName);
                File.Move(quellDateiPfad, Nutzer_ArchivPfad);
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void DokumenteEinblenden()
        {
            
            //Kein Sprachen abgleich möglich alle Sichtbar machen
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
            // Dateien aus dem Quellordner in den Zielordner kopieren
            foreach (string datei in Directory.GetFiles(GlobalVariables.Pfad_QuelleVorlagen))
            {
                string dateiName = System.IO.Path.GetFileName(datei);
                string zielDateiPfad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, dateiName);

                // Datei kopieren, falls sie nicht schon existiert
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
            // Zielordner erstellen, falls sie nicht existieren
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
            int row_count = DB_Serviceaufträge.Rows.Count;

            for (int i = 0; i < row_count; i++)
            {
                string DB_AuftragsNR = Convert.ToString(DB_Serviceaufträge.Rows[i][1]);

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
                }
            }
        }
        // Unterfunktion btn_AuftragsNR_suchen_Click
        private void PfadeFestlegen()
        {
            string basisOrdner = AppDomain.CurrentDomain.BaseDirectory;
            string relativerPfad = @"Vorlagen";
            GlobalVariables.Pfad_QuelleVorlagen = System.IO.Path.Combine(basisOrdner, relativerPfad);

            //Das Ergebniss der Vorran gegangenen Pings an Fileserver und DB prüfen
            if (GlobalVariables.Online_or_Offline)
            {
                //Online Pfade setzen
                GlobalVariables.Pfad_AuftragsOrdner = string.Format(Properties.Resources.Pfad_AuftragsOrdner_On, GlobalVariables.AuftragsNR);
                GlobalVariables.Pfad_Anhaenge = string.Format(Properties.Resources.Pfad_Anhaenge_On, GlobalVariables.AuftragsNR);
                GlobalVariables.Pfad_Unterschriften = string.Format(Properties.Resources.Pfad_Signatures_On, GlobalVariables.AuftragsNR);
            }
            else
            {
                //Offline Pfade setzen
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
            if(tb_Land_Startseite.Text == "" && tb_KundenFirmenName_Startseite.Text == "" && dp_AbDatum_Startseite.Text == "")
            {
                MessageBox.Show("Bitte geben Sie mindestens ein Suchkriterium ein.");
                return;
            }

            var view = dataGrid.ItemsSource as DataView;
            DataTable table = view?.ToTable(); // enthält alle gefilterten und sortierten Zeilen

            var filteredData = table?.AsEnumerable(); // als IEnumerable<DataRow>

            if (!string.IsNullOrEmpty(tb_Land_Startseite.Text))
            {
                filteredData = filteredData.Where(row => row.Field<string>("Land").ToUpper().Contains(tb_Land_Startseite.Text.ToUpper()));
            }

            if (!string.IsNullOrEmpty(tb_KundenFirmenName_Startseite.Text))
            {

                filteredData = filteredData.Where(row => row.Field<string>("A0Name1").ToUpper().Contains(tb_KundenFirmenName_Startseite.Text.ToUpper()));
            }

            if (dp_AbDatum_Startseite.SelectedDate.HasValue)
            {
                if(dp_AbDatum_Startseite.SelectedDate.Value <= DateTime.Today) 
                { 
                    var abDatum = dp_AbDatum_Startseite.SelectedDate.Value;
                    filteredData = filteredData.Where(row => row.Field<DateTime>("Belegdatum") >= abDatum);
                }
                else
                {
                    MessageBox.Show("Das Datum darf nicht in der Zukunft liegen.");
                    return;
                }
            }

            if (filteredData == null || !filteredData.Any())
            {
                dataGrid.BorderBrush = Brushes.Red;
                dataGrid.BorderThickness = new Thickness(4);
                MessageBox.Show("Keine Daten gefunden.");
                dataGrid.BorderBrush = Brushes.Black;
                dataGrid.BorderThickness = new Thickness(1);
                return;
            }

            try
            {
                dataGrid.ItemsSource = filteredData.CopyToDataTable().DefaultView;
            }
            catch (InvalidOperationException)
            {
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
            // Setzt den Cursor in die TextBox
            tb_Auftragsnummer.Focus();

            // Text markieren
            tb_Auftragsnummer.SelectAll();
        }

        private void AuftragsNrSucheStarten(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btn_AuftragsNr_suchen_Click(sender, e);
            }
        }

        private void StarteTabellenFilter(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                OnSearchClick(sender, e);
            }
        }
    
        private void ClearGlobalVariables()
        {
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
