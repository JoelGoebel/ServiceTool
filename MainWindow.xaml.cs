using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.DesignerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.TextFormatting;
using System.Windows.Threading;
using static ServiceTool.MainWindow;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

//TODO-ListIn




namespace ServiceTool
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public bool _isInitialized = false;
        private bool _blockiereUControlWechsel = false;
        bool isFirstLoad = true;
        public MainWindow()
        {
            InitializeComponent();
            //Set the size of the MainWindow to 95% of the screen size
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;

            this.Width = screenWidth * 0.95;
            this.Height = screenHeight * 0.95;

            //Sets Variable to true if MainWindow is loaded
            this.Loaded += MainWindow_Loaded;

            //Saves all the Cell Mappings in the GlobalVariables Dictionarys
            SaveCellMapping_InDictionarys();

            //Variables for Language switch
            List<string> Lbl_Names = new List<string>();
            List<string> Lbl_Content_German = new List<string>();
            List<string> Lbl_Content_English = new List<string>();

            //Get the content of the Labels from the JSON file
            GetLabelContent();

            //Test if the server is reachable
            bool File_Connection_Test = IstServerErreichbar(Properties.Resources.IP_File02);
            bool DB_Connection_Test = IstServerErreichbar(Properties.Resources.IP_SQL04);

            if (File_Connection_Test && DB_Connection_Test)
            {
                //Funktion to execute the SQL Query and collect the data from the database
                collect_Data_From_Database();
                //Set label to Online so that the user knows that the program is online
                GlobalVariables.Online_or_Offline = true;
                lbl_OnlineOfflineAnzeige.Content = "Online";
                lbl_OnlineOfflineAnzeige.Background = Brushes.Green;
            }
            else
            {
                //Set label to Offline so that the user knows that the program is offline
                GlobalVariables.Online_or_Offline = false;
                lbl_OnlineOfflineAnzeige.Content = "Offline";
                lbl_OnlineOfflineAnzeige.Background = Brushes.Red;
            }
            // Place the Startseite in the Content Control
            CC.Content = new Startseite();
        }


        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _isInitialized = true;
        }

        SprachTabelle sprachtabelle_IBNP = new SprachTabelle();
        SprachTabelle sprachtabelle_IBNP_MRS = new SprachTabelle();
        public void GetLabelContent()
        {
            string UserPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            //Read all Data From JSON
            string json = File.ReadAllText(Properties.Resources.Path_LanguageJson_IBNP);
            //Convert the string in a List of a selfmade class SprachtabelleEntry
            List<SprachtabelleEntry> sprachtabelleEntries = JsonConvert.DeserializeObject<List<SprachtabelleEntry>>(json);

            //Add all entrys to Languagetable IBNP and IBNP_MRS
            foreach (SprachtabelleEntry entry in sprachtabelleEntries)
            {
                sprachtabelle_IBNP.Lbl_Names.Add(entry.LBL_Names);
                sprachtabelle_IBNP.Lbl_Content_German.Add(entry.Deutsch);
                sprachtabelle_IBNP.Lbl_Content_English.Add(entry.Englisch);
            }

            json = File.ReadAllText(Properties.Resources.Path_LanguageJson_IBNP_MRS);
            List<SprachtabelleEntry> sprachtabelleEntries_MRS = JsonConvert.DeserializeObject<List<SprachtabelleEntry>>(json);

            foreach (SprachtabelleEntry entry in sprachtabelleEntries_MRS)
            {
                sprachtabelle_IBNP_MRS.Lbl_Names.Add(entry.LBL_Names);
                sprachtabelle_IBNP_MRS.Lbl_Content_German.Add(entry.Deutsch);
                sprachtabelle_IBNP_MRS.Lbl_Content_English.Add(entry.Englisch);
            }

        }
        public void SetLanguage(string Seite)
        {
            object Dokument = null;
            SprachTabelle sprachtabelle = new SprachTabelle();
            //Load the correct Document based on the selected page
            switch (Seite)
            {
                case "IbnP":
                    Dokument = CC.Content as Inbetriebnahme_Protokoll;
                    sprachtabelle = sprachtabelle_IBNP;
                    break;

                case "Serviceanforderungen":
                    Dokument = CC.Content as Service_Anforderung;
                    
                    break;

                case "Stundennachweis":
                    Dokument = CC.Content as Stundennachweis;
                   
                    break;

                case "Interner_Bericht":
                    Dokument = CC.Content as Interner_Bericht;
                    break;

                case "IbnP_MRS":
                    Dokument = CC.Content as Inbetriebnahmeprotokoll_MRS;
                    sprachtabelle = sprachtabelle_IBNP_MRS;
                    break;
            }
            // Set the language of the labels in the document based on the selected language
            if (Dokument is FrameworkElement element) { 
                foreach (string lblName in sprachtabelle.Lbl_Names)
                {
                    // Find the label by its name in the document
                    Label lbl = (Label)element.FindName(lblName);

                    if (GlobalVariables.Sprache_Kunde == "D")
                    {
                        lbl.Content = sprachtabelle.Lbl_Content_German[sprachtabelle.Lbl_Names.IndexOf(lblName)];
                    }
                    else
                    {
                        lbl.Content = sprachtabelle.Lbl_Content_English[sprachtabelle.Lbl_Names.IndexOf(lblName)];
                    }
                    
                }
            }
        }//ENde Set Language
        public void SaveCellMapping_InDictionarys()
        {
            //Read JSON Files and save the Cell Mappings in the GlobalVariables Dictionarys
            //Cell Mapping for Inbetriebnahme Protokoll
            string json = File.ReadAllText(Properties.Resources.Path_CellMappingIBNP);
            var cellMappings = JsonConvert.DeserializeObject<List<CellMapping>>(json);
            GlobalVariables.CellMapping_IbnP = cellMappings.ToDictionary(cm => cm.Zelle, cm => cm.Feldname);

            json = File.ReadAllText(Properties.Resources.Path_CellMappingIBNP_MRS);
            cellMappings = JsonConvert.DeserializeObject<List<CellMapping>>(json);
            GlobalVariables.CellMapping_IBNP_MRS = cellMappings.ToDictionary(cm => cm.Zelle, cm => cm.Feldname);

            json = File.ReadAllText(Properties.Resources.Path_CellMappingServiceAnforderungen);
            cellMappings = JsonConvert.DeserializeObject<List<CellMapping>>(json);
            GlobalVariables.CellMapping_ServiceAnforderungen = cellMappings.ToDictionary(cm => cm.Zelle, cm => cm.Feldname);

            json = File.ReadAllText(Properties.Resources.Path_CellMappingStundenachweis);
            cellMappings = JsonConvert.DeserializeObject<List<CellMapping>>(json);
            GlobalVariables.CellMapping_Stundenachweis = cellMappings.ToDictionary(cm => cm.Zelle, cm => cm.Feldname);            

            json = File.ReadAllText(Properties.Resources.Path_CellMapping_InternerBericht);
            cellMappings = JsonConvert.DeserializeObject<List<CellMapping>>(json);
            GlobalVariables.CellMapping_InternerBericht = cellMappings.ToDictionary(cm => cm.Zelle, cm => cm.Feldname);
        }
        public static bool IstServerErreichbar(string serverAdresse, int timeout = 1000)
        {
            //Check if the server is reachable by sending a ping request
            try
            {
                using (Ping pingSender = new Ping())
                {
                    PingReply antwort = pingSender.Send(serverAdresse, timeout);
                    return antwort.Status == IPStatus.Success;
                }
            }
            catch (PingException)
            {
                // Behandlung von Ping-spezifischen Ausnahmen
                return false;
            }
            catch (Exception)
            {
                // Behandlung anderer Ausnahmen
                return false;
            }
        }
        public void collect_Data_From_Database()
        {
            //Set the connection string from the resources
            string Connectionstring = Properties.Resources.Connectionstring;
            //Sét the SQL Query from the resources
            string DB_Query = Properties.Resources.DB_Abfrage;

            //Execute the SQL Query and fill the DataTable with the data from the database
            using (SqlConnection connection = new SqlConnection(Connectionstring)) 
            {
                SqlDataAdapter adapter = new SqlDataAdapter(DB_Query, connection);
                GlobalVariables.dt = new DataTable();
                adapter.Fill(GlobalVariables.dt);
            }

            //Write the Data in the Console to make the Deugging easier 
            // Ausgabe der Spaltennamen
            //foreach (DataColumn column in GlobalVariables.dt.Columns)
            //{
            //    Console.Write($"{column.ColumnName}\t");
            //}
            //Console.WriteLine();

            // Ausgabe der Zeilen
            //foreach (DataRow row in GlobalVariables.dt.Rows)
            //{
            //    foreach (var item in row.ItemArray)
            //    {
            //        Console.Write($"{item}\t");
            //    }
            //    Console.WriteLine();
            //}
        }

        private void rbt_Startseite_Checked(object sender, RoutedEventArgs e) // wenn der  radiobutton von der Startseite angehackt wird, wird die Startseite erstellt und im Content Control platziert
        {
            //generate new Startseite if Radiobutton is Collected ( Radiobutton is already checked after starting the Programm)
            if (_blockiereUControlWechsel) return;
            CC.Content = new Startseite();

        }

        //Funktion starts after the Radiobutton for Serviceanforderungen is checked
        private void rbt_ServiceAnforderung_Checked(object sender, RoutedEventArgs e)
        {
            //Create New Document (ServiceAnforderugen) and place it in the Content Control
            var sa = new Service_Anforderung();
            CC.Content = sa;
            
            
            string Auftragsnummer = GlobalVariables.AuftragsNR;                      

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner,"Service_Anforderungen.xlsx");

            //load Saved Data from the Excel file into the Service_Anforderung Document
            Laden(ExcelFilePath, "Serviceanforderungen");

            sa.tb_Auftragsnummer.Text = GlobalVariables.AuftragsNR;

            //Check íf Auftrag is in DB because the Informations are not in the Variables if not 
            if (GlobalVariables.auftraginDB == true)
            {
                sa.tb_Anschrift_1_Anforderung.Text = GlobalVariables.Anschrift_1;
                sa.tb_Anschrift_2_Anforderung.Text= GlobalVariables.Anschrift_2;
                sa.tb_Kunde_Anforderung.Text = GlobalVariables.Kunde;
                sa.tb_KundenNr.Text = GlobalVariables.KundenNummer;
                sa.tb_Land.Text = GlobalVariables.Land;
            }
            GlobalVariables.Land = sa.tb_Land.Text;

        }
        //Funktion gets triggerd after unchecking Radiobutton Serviceanforderungen
        private void rbt_ServiceAnforderung_UnChecked(object sender, RoutedEventArgs e) 
        {
            var sa = CC.Content as Service_Anforderung;
            // Proceed only if the control has been properly initialized
            if (_isInitialized)
            {

                if (sa is IValidierbar validierbar)
                {
                    // If required fields are missing, show an error message
                    if (validierbar.HatFehlendePflichtfelder(out string fehlermeldung))
                    {
                        MessageBox.Show(fehlermeldung, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                        // Prevent switching to another UserControl temporarily
                        _blockiereUControlWechsel = true;
                        // Schedule re-checking the radio button after the current operation finishes
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            rbt_ServiceAnforderung.IsChecked = true;
                            _blockiereUControlWechsel = false;
                        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

                        return;
                    }
                }
            }


            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Service_Anforderungen.xlsx");
            // Call the save function with the file path and worksheet name
            speichern(ExcelFilePath, "Serviceanforderungen");

            // Transfer form data into global variables for later use
            GlobalVariables.Kunde = sa.tb_End_Kunde.Text;
            GlobalVariables.Ansprechpartner = sa.tb_Ansprechpartner_Anforderung.Text;
            GlobalVariables.Anschrift_1 = sa.tb_Anschrift_1_Anforderung.Text;
            GlobalVariables.Anschrift_2 = sa.tb_Anschrift_2_Anforderung.Text;
            GlobalVariables.Anreise = sa.cb_Anreise.Text;
            GlobalVariables.ServiceTechnicker = sa.tb_Servicetechniker_Anforderung.Text;

            // Only store machine types if the dropdown text is not just a space
            if (sa.cb_Maschinentyp_1.Text != " ")
            {
                GlobalVariables.Maschiene_1 = sa.cb_Maschinentyp_1.Text;
            }
            if (sa.cb_Maschinentyp_2.Text != " ")
            {
                GlobalVariables.Maschiene_2 = sa.cb_Maschinentyp_2.Text;
            }
            if (sa.cb_Maschinentyp_3.Text != " ")
            {
                GlobalVariables.Maschiene_3 = sa.cb_Maschinentyp_3.Text;
            }
            if (sa.cb_Maschinentyp_4.Text != " ")
            {
                GlobalVariables.Maschiene_4 = sa.cb_Maschinentyp_4.Text;
            }

            //Save the machine size and machine number in the global variables
            GlobalVariables.Baugroeße_1 = sa.cb_BauGröße_1.Text;
            GlobalVariables.Baugroeße_2 = sa.cb_BauGröße_2.Text;
            GlobalVariables.Baugroeße_3 = sa.cb_BauGröße_3.Text;
            GlobalVariables.Baugroeße_4 = sa.cb_BauGröße_4.Text;

            GlobalVariables.MaschinenNr_1 = sa.tb_MaschNr_1.Text;
            GlobalVariables.MaschinenNr_2 = sa.tb_MaschNr_2.Text;
            GlobalVariables.MaschinenNr_3 = sa.tb_MaschNr_3.Text;
            GlobalVariables.MaschinenNr_4 = sa.tb_MaschNr_4.Text;

            GlobalVariables.Land = sa.tb_Land.Text;

            GlobalVariables.Material = sa.tb_Material.Text;

            // If both start and end dates for the visit are selected, save them
            if (sa.dp_Besuchsdatum_Start.SelectedDate != null && sa.dp_Besuchsdatum_Ende.SelectedDate != null)
            {
                GlobalVariables.StartServiceEinsatz = (DateTime)sa.dp_Besuchsdatum_Start.SelectedDate;
                GlobalVariables.EndeServiceEinsatz = (DateTime)sa.dp_Besuchsdatum_Ende.SelectedDate;
            }
        }

        private ExportDataWorkingTime GetDataForExport(Stundennachweis sn, string Day)
        {
            ExportDataWorkingTime exportData = new ExportDataWorkingTime();
            
            exportData.EinsatzDatum_Start = GlobalVariables.StartServiceEinsatz;
            exportData.EinsatzDatum_Ende = GlobalVariables.EndeServiceEinsatz;
            exportData.Auftragsnummer = GlobalVariables.AuftragsNR;
            exportData.ServiceTechnicker = sn.tb_Servicetechiker_Stunden.Text;
            switch (Day)
            {
                case "Mo":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Mo_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Mo_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Mo_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Mo_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Mo_Stunden.Text);
                    break;

                case "Di":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Di_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Di_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Di_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Di_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Di_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Di_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Di_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Di_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Di_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Di_Stunden.Text);
                    break;

                case "Mi":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Mi_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Mi_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Mi_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Mi_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Mi_Stunden.Text);
                    break;

                case "Do":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Do_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Do_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Do_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Do_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Do_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Do_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Do_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Do_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Do_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Do_Stunden.Text);
                    break;

                case "Fr":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Fr_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Fr_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Fr_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Fr_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Fr_Stunden.Text);
                    break;

                case "Sa":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_Sa_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_Sa_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_Sa_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_Sa_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_Sa_Stunden.Text);
                    break;

                case "So":
                    exportData.ArbeitsZeit_Start = TimeSpan.Parse(sn.cb_Von_So_Stunden.Text);
                    exportData.ArbeitsZeit_Ende = TimeSpan.Parse(sn.cb_Bis_So_Stunden.Text);
                    exportData.ArbeitsZeit_Pause = TimeSpan.Parse(sn.cb_Pause_So_Stunden.Text);
                    exportData.ArbeitsZeit_Start_S2 = TimeSpan.Parse(sn.cb_Von_So_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Ende_S2 = TimeSpan.Parse(sn.cb_Bis_So_S2_Stunden.Text);
                    exportData.ArbeitsZeit_Pause_S2 = TimeSpan.Parse(sn.cb_Pause_So_S2_Stunden.Text);
                    exportData.ArbeitsZeit_NormalHours = TimeSpan.Parse(sn.tb_NormalStd_So_Stunden.Text);
                    exportData.ArbeitsZeit_Overtime = TimeSpan.Parse(sn.tb_Ueberstunden_So_Stunden.Text);
                    exportData.ArbeitsZeit_NightWork = TimeSpan.Parse(sn.tb_Nachtarbeit_So_Stunden.Text);
                    exportData.ArbeitsZeit_Gesamt = TimeSpan.Parse(sn.tb_GesamtStd_So_Stunden.Text);
                    break;
            }

            exportData.AnreiseDatum_Start = (DateTime)sn.dp_AnreiseDatum_Stunden.SelectedDate;
            exportData.AnreiseDatum_Ende = (DateTime)sn.dp_AnreiseDatumAnkunft_Stunden.SelectedDate;
            exportData.Anreise_Startzeit = TimeSpan.Parse(sn.cb_Anreise_Fahrtbeginn_Stunden.Text);
            exportData.Anreise_Endezeit = TimeSpan.Parse(sn.cb_Anreise_Fahrtende_Stunden.Text);
            exportData.Anreise_Pause = TimeSpan.Parse(sn.cb_Anreise_Pause_Stunden.Text);
            exportData.Anreise_DauerGesamt = TimeSpan.Parse(sn.tb_Anreisedauer_Gesamt_Stunden.Text);
            exportData.Anreise_KM = sn.tb_AnreiseWeg_Stunden.Text;
            exportData.AbreiseDatum_Start = (DateTime)sn.dp_AbreiseDatum_Stunden.SelectedDate;
            exportData.AbreiseDatum_Ende = (DateTime)sn.dp_AbreiseDatumAnkunft_Stunden.SelectedDate;
            exportData.Abreise_Startzeit = TimeSpan.Parse(sn.cb_Abreise_Fahrtbeginn_Stunden.Text);
            exportData.Abreise_Endezeit = TimeSpan.Parse(sn.cb_Abreise_Fahrtende_Stunden.Text);
            exportData.Abreise_Pause = TimeSpan.Parse(sn.cb_Abreise_Pause_Stunden.Text);
            exportData.Abreise_DauerGesamt = TimeSpan.Parse(sn.tb_Abreisedauer_Gesamt_Stunden.Text);
            exportData.Abreise_KM = sn.tb_AbreiseWeg_Stunden.Text;

            return exportData;
        }
        private void rbt_Stundennachweis_Checked(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;
            // Hier wird der Stundennachweis geladen
            var sn = new Stundennachweis();
            CC.Content = sn;            

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm");

            // Prüfen, ob die Datei am angegebenen Pfad existiert
            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Benutze das ExcelPackage, um die Excel-Datei zu öffnen. Der using-Block stellt sicher, dass die Datei geschlossen wird, wenn der Block beendet ist
                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu

                    Laden(ExcelFilePath, "Stundennachweis");
                    //Textboxen die aus den Service anforderungen übernommen werden
                    sn.tb_Servicetechiker_Stunden.Text = GlobalVariables.ServiceTechnicker;
                    sn.tb_Servicetechiker_Stunden.Focusable = false;
                    sn.tb_Kunde_Stunden.Text = GlobalVariables.Kunde;
                    sn.tb_Kunde_Stunden.Focusable = false;
                    sn.tb_Ansprechpartner_Stunden.Text = GlobalVariables.Ansprechpartner;
                    sn.tb_Ansprechpartner_Stunden.Focusable = false;
                    sn.tb_Anschrift_1_Stunden.Text = GlobalVariables.Anschrift_1;
                    sn.tb_Anschrift_1_Stunden.Focusable = false;
                    sn.tb_Anschrift_2_Stunden.Text = GlobalVariables.Anschrift_2;
                    sn.tb_Anschrift_2_Stunden.Focusable = false;
                    if(GlobalVariables.Anreise != "" && GlobalVariables.Anreise !="-bitte auswählen-")
                    {
                        sn.cb_Verkehrsmittel_Stunden.Text = GlobalVariables.Anreise;
                        sn.cb_Verkehrsmittel_Stunden.Focusable = false;
                    }               
                }
            }
            else
            {
                MessageBox.Show("Die Excel-Datei wurde nicht gefunden. Oder es wurde keine Auftragsnummer eingegeben", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void Stundennachweis_Speichern() // Funktion zum sichern der Werte des Programms in der Excel datei des Auftrags
        {
            if (_blockiereUControlWechsel) return;
            var sn = CC.Content as Stundennachweis;
            string ExcelFilePath = "";

            //Set Correct Path depending on the selected week in the ComboBox
            switch (sn.cb_Siteswitch_Stunden.Text)
            {
                case "Woche 1":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm ");
                    break;
                case "Woche 2":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_2.xlsm ");
                    break;
                case "Woche 3":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_3.xlsm ");
                    break;
                case "Woche 4":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_4.xlsm ");
                    break;
                case "Woche 5":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_5.xlsm ");
                    break;
                case "Woche 6":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_6.xlsm ");
                    break;
                case "Woche 7":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_7.xlsm ");
                    break;
                case "Woche 8":
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_8.xlsm ");
                    break;
            }

            for (int i = 0; i < 7; i++)
            {

                string day = "";
                switch (i)
                {
                    case 0:
                        day = "Mo";
                        break;
                    case 1:
                        day = "Di";
                        break;
                    case 2:
                        day = "Mi";
                        break;
                    case 3:
                        day = "Do";
                        break;
                    case 4:
                        day = "Fr";
                        break;
                    case 5:
                        day = "Sa";
                        break;
                    case 6:
                        day = "So";
                        break;
                }
                ExportDataWorkingTime exportData = GetDataForExport(sn, day);
                //TODO HIER DENN EXPORT EINBAUEN VARIABLEN ANSPRECHBAR MIT exportData.(SpaltenName)
            }

            speichern(ExcelFilePath, "Stundennachweis");

            // Ab hier ist die Funktion nur zum Speichern der Signaturen
            

            // Prüfen, ob die Datei am angegebenen Pfad existiert
            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Benutze das ExcelPackage, um die Excel-Datei zu öffnen. Der using-Block stellt sicher, dass die Datei geschlossen wird, wenn der Block beendet ist
                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu

                    string imagepath_sign_technican = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureemployee.png");
                    Dispatcher.Invoke(() =>
                    {
                        if (!File.Exists(imagepath_sign_technican) && sn.ic_Unterschrift_Technicker.Strokes.Count != 0)
                        { 
                            SaveSignatureAsImage(sn.ic_Unterschrift_Technicker, imagepath_sign_technican);
                        }
                    }, DispatcherPriority.Render);

                    
                    string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureCustomer.png");
                    Dispatcher.Invoke(() =>
                    {
                        if (!File.Exists(ImagePath_Sign_Kunde) && sn.ic_UnterschriftKunde_Stunden.Strokes.Count != 0)
                        {
                            SaveSignatureAsImage(sn.ic_UnterschriftKunde_Stunden, ImagePath_Sign_Kunde);
                        }
                    }, DispatcherPriority.Render);

                    package.Save();
                    package.Dispose();
                }
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;

                try
                {
                    string imagepath_sign_technican = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureemployee.png");
                    string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureCustomer.png");

                    excelApp.Visible = false;
                    workbook = excelApp.Workbooks.Open(ExcelFilePath);

                    // Das Makro ausführen
                    if (File.Exists(ImagePath_Sign_Kunde) && File.Exists(imagepath_sign_technican))
                    {
                        excelApp.Run("Signaturen_einfügen");
                    }

                    // Speichern und Schließen
                    workbook.Save();
                    workbook.Close();
                }
                catch (Exception ex)
                {
                    // Fehlerbehandlung
                    Console.WriteLine($"Fehler beim Ausführen des Makros: {ex.Message}");
                }
                finally
                {
                    // Beende die Excel-Anwendung
                    excelApp.Quit();

                    // Freigeben von COM-Objekten
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    workbook = null;
                    excelApp = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }
        private void rbt_Stundenachweis_UnChecked(object sende, RoutedEventArgs e) { Stundennachweis_Speichern(); } // Trigger für die Speicher Funktion (ausgelöst wenn Radiobutton nicht merh angehackt ist)

        private void rbt_InternerBericht_Checked(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;
            var ib = new Interner_Bericht();
            CC.Content = ib;

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht.xlsx");

            // Prüfen, ob die Datei am angegebenen Pfad existiert
            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Benutze das ExcelPackage, um die Excel-Datei zu öffnen. Der using-Block stellt sicher, dass die Datei geschlossen wird, wenn der Block beendet ist

                Laden(ExcelFilePath, "Interner_Bericht");
            }
            //TODO Hier noch so machen das CB_Einheit_M den Wert von der Exceldatei bekommt
            if (ib.CB_Einheit_M.Text != "")
            {
                ib.CB_Einheit_T1.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_B1.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_T2.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_T3.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_B2.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_B3.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_T4.Text = ib.CB_Einheit_M.Text; 
                ib.CB_Einheit_B4.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_TB0.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_D0.Text = ib.CB_Einheit_M.Text;
                ib.CB_Einheit_Sonstige.Text = ib.CB_Einheit_M.Text;
            }
            

        } // Lade Funktion für den internen Bericht
        private void rbt_InternerBericht_UnChecked(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner,"interner_Bericht.xlsx");

            speichern(ExcelFilePath, "Interner_Bericht");
        }// Speicher Funktion für den internen Bericht

        private void rbt_InbetriebnahmeProtokoll_Checked(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;
            InbetriebnahmeProtokoll_Laden(GlobalVariables.SelectedItemIbnP);
        }// Funktion zum Laden der Infos aus der Excel-Datei des Auftrags (ausgelöst wenn Radiobutton angehackt wird)
        public void InbetriebnahmeProtokoll_Laden(string selectedItem)
        {
            //Erstelle neues Inbetriebnahme Protokoll
            var ibnP = new Inbetriebnahme_Protokoll(isFirstLoad);
            CC.Content = ibnP;//Setze das IbnP in den Content Control im MainWindow

            //Sprache nach auftrag setzen
            SetLanguage("IbnP");

            string ExcelFilePath = "";                 
            
            if(GlobalVariables.Maschiene_1 !="" && GlobalVariables.Maschiene_1 != "MRS" && GlobalVariables.Maschiene_1 != "Jump")
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                Laden(ExcelFilePath, "IbnP");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_1;
            }
            else if(GlobalVariables.Maschiene_2 != "" && GlobalVariables.Maschiene_2 != "MRS" && GlobalVariables.Maschiene_2 != "Jump")
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_2.xlsm");
                Laden(ExcelFilePath, "IbnP");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_2;
            }
            else if (GlobalVariables.Maschiene_3 != "" && GlobalVariables.Maschiene_3 != "MRS" && GlobalVariables.Maschiene_3 != "Jump")
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                Laden(ExcelFilePath, "IbnP");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_3;
            }
            else if (GlobalVariables.Maschiene_4 != "" && GlobalVariables.Maschiene_4 != "MRS" && GlobalVariables.Maschiene_4 != "Jump")
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                Laden(ExcelFilePath, "IbnP");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_4;
            }

            ibnP.tb_Kunde_ibnProtokoll.Text = GlobalVariables.Kunde;
            ibnP.tb_Ansprechpartner_ibnProtokoll.Text = GlobalVariables.Ansprechpartner;
            ibnP.tb_KundeMaterial_ibnProtokoll.Text = GlobalVariables.Material;
            isFirstLoad = false;
        }//Ende InbP Laden        
                                 
        public void InbetriebnahmeProtokoll_Speichern(string lastSelectedSite)
        {
            if (_blockiereUControlWechsel) return;
            var ibnP = CC.Content as Inbetriebnahme_Protokoll;

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ImagePath_Sign_Kunde = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner,"Anhaenge\\Unterschriften\\ibnPSignatureCustomer.png");

            string ImagePath_Sign_Technican = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee.png");

            string ExcelFilePath = "";

            // da es die möglichkeit mehrer IbnP gibt muss überprüft werden welche aktuell bearbeitet wurde an dem Punkt wo der  IbnP Radiobutton abgehackt wurde
            if (lastSelectedSite == "" || lastSelectedSite == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                speichern(ExcelFilePath, "IbnP");
                ImagePath_Sign_Kunde = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer.png");
                ImagePath_Sign_Technican = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee.png");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_1;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_2.xlsm");
                speichern(ExcelFilePath, "IbnP");
                ImagePath_Sign_Kunde = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_2.png");
                ImagePath_Sign_Technican = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_2.png");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_2;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                speichern(ExcelFilePath, "IbnP");
                ImagePath_Sign_Kunde = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_3.png");
                ImagePath_Sign_Technican = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_3.png");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_3;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                speichern(ExcelFilePath, "IbnP");
                ImagePath_Sign_Kunde = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_4.png");
                ImagePath_Sign_Technican = Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_4.png");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_4;
            }           

            // Ab hier  allles zum speichern der Signatur

            // Prüfen, ob die Datei am angegebenen Pfad existiert
            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                // Benutze das ExcelPackage, um die Excel-Datei zu öffnen. Der using-Block stellt sicher, dass die Datei geschlossen wird, wenn der Block beendet ist
                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu
                    

                    if (ibnP.ic_Unterschrift_Kunde_ibnProtokoll.Strokes.Count != 0 && ibnP.ic_Unterschrift_Servicetechniker_ibnProtokoll.Strokes.Count != 0 && worksheet.Cells["H55"].Text == "Nein")
                    {

                        if (!File.Exists(ImagePath_Sign_Kunde)) { SaveSignatureAsImage(ibnP.ic_Unterschrift_Kunde_ibnProtokoll, ImagePath_Sign_Kunde); }

                        if (!File.Exists(ImagePath_Sign_Technican)) { SaveSignatureAsImage(ibnP.ic_Unterschrift_Servicetechniker_ibnProtokoll, ImagePath_Sign_Technican); }



                    }
                    package.Save();
                    package.Dispose();
                }
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;

                try
                {
                    excelApp.Visible = false;
                    workbook = excelApp.Workbooks.Open(ExcelFilePath);
                   
                    // Das Makro ausführen um die PNG der Signaturen an der Richtigen stelle der Excel datei einzufügen
                    excelApp.Run("Signaturen_einfügen");
             
                    // Speichern und Schließen
                    workbook.Save();
                    workbook.Close();
                }
                catch (Exception ex)
                {
                    // Fehlerbehandlung
                    Console.WriteLine($"Fehler beim Ausführen des Makros: {ex.Message}");
                }
                finally
                {
                    // Beende die Excel-Anwendung
                    excelApp.Quit();

                    // Freigeben von COM-Objekten
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                    workbook = null;
                    excelApp = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }

        } //Ende IbnP Speichern MW 
        private void rbt_InbetriebnahmeProtokoll_UnChecked(object sender, RoutedEventArgs e) 
        { 
            string lastSelectedSite = GlobalVariables.LastSelectedSiteIbnP;
            InbetriebnahmeProtokoll_Speichern(lastSelectedSite); 
        } //Trigger für die speicher Funktion

        private void InbetriebnahmeProtokoll_MRS_Laden(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;

            var IbnP_MRS = new Inbetriebnahmeprotokoll_MRS();
            CC.Content = IbnP_MRS;

            SetLanguage("IbnP_MRS");

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");

            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];                   

                    Laden(ExcelFilePath, "IbnP_MRS");
                    IbnP_MRS.tb_Kunde_ibnProtokoll_MRS.Text = GlobalVariables.Kunde;
                }
            }
            
            IbnP_MRS.tb_Kunde_ibnProtokoll_MRS.Text = GlobalVariables.Kunde;
            IbnP_MRS.tb_ExtruderTyp_ibnProtokoll_MRS.Text = GlobalVariables.Maschiene_1 + GlobalVariables.Baugroeße_1;
            IbnP_MRS.tb_Seriennummer_ibnProtokoll_MRS.Text = GlobalVariables.MaschinenNr_1;

        }
        private void IbnP_MRS_Speichern(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;
            var IbnP_MRS = CC.Content as Inbetriebnahmeprotokoll_MRS;

            string Auftragsnummer = GlobalVariables.AuftragsNR;
            string LastSelectedItem = GlobalVariables.LastSelectedItem_MRS; // Letztes ausgewähltes Item für die MRS Inbetriebnahme Protokolle
            string ExcelFilePathSave = "";
            string ImagePath_Sign_Kunde = "";
            string ImagePath_Sign_Technican = "";


            //Set ExcelfilePath for saving last Selected Site
            if (LastSelectedItem == "" || LastSelectedItem == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");
                speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS.png");
                if(!File.Exists(ImagePath_Sign_Kunde))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                if (!File.Exists(ImagePath_Sign_Technican)) SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_2.xlsx");
                speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_2.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_2.png");
                if(!File.Exists(ImagePath_Sign_Kunde))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                if(!File.Exists(ImagePath_Sign_Technican))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_3.xlsx");
                speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_3.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_3.png");
                if (!File.Exists(ImagePath_Sign_Kunde)) SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                if(!File.Exists(ImagePath_Sign_Technican))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }
            else if (LastSelectedItem == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_4.xlsx");
                speichern(ExcelFilePathSave, "IbnP_MRS");
                ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureCustomer_MRS_4.png");
                ImagePath_Sign_Technican = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Anhaenge\\Unterschriften\\ibnPSignatureEmployee_MRS_4.png");
                if(!File.Exists(ImagePath_Sign_Kunde))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Kunde_ibnP_MRS, ImagePath_Sign_Kunde);
                if(!File.Exists(ImagePath_Sign_Technican))SaveSignatureAsImage(IbnP_MRS.ic_Unterschrift_Servicetechniker_ibnP_MRS, ImagePath_Sign_Technican);
            }

            
        }

        //Funktion um einmal alle Seiten zu laden
        public void CheckAllRadioButtons()
        {
            rbt_ServiceAnforderung.IsChecked = true;

            rbt_Startseite.IsChecked = true; // Zurück zur Startseite
        } // wird beim Start aus geführt um gewisse werte und Objecte vorzuladen damit in anderen Funktionen mit diesen gearbeitet werden kann
        private void Ordner_oeffnen_Anhaenge(object sender, RoutedEventArgs e)
        {
            // Save the path for attachments from GlobalVariables
            string Pfad_fuerAnhaenge = GlobalVariables.Pfad_Anhaenge;

            Process.Start("explorer.exe", Pfad_fuerAnhaenge);
        } // Funktion um mit einem Button klick in den Anhang ordner des Auftrags zu gelangen

        public void SaveSignatureAsImage(InkCanvas inkCanvas, string filePath)
        {
            //Layout aktualisieren, damit Größen und Striche definitiv bereitstehen
            inkCanvas.UpdateLayout();  // sicherstellen, dass ActualWidth/Height korrekt

            int width = (int)inkCanvas.ActualWidth;
            int height = (int)inkCanvas.ActualHeight;
            if (width == 0 || height == 0) return; // InkCanvas nicht sichtbar oder keine Größe

            //DrawingVisual erzeugen und darin das InkCanvas "nachmalen"
            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                if (inkCanvas.Background != null)
                {
                    // Hintergrund als Brush füllen (z.B. Farbe) über die ganze Fläche
                    dc.DrawRectangle(inkCanvas.Background, null, new Rect(0, 0, width, height));
                }
                // Alle Striche zeichnen
                foreach (System.Windows.Ink.Stroke stroke in inkCanvas.Strokes)
                {
                    stroke.Draw(dc);  // Stroke zeichnet sich selbst mit seinen DrawingAttributes
                }
                
            } // DrawingContext auto-close here

            //RenderTargetBitmap mit passendem PixelFormat anlegen und DrawingVisual rendern
            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);

            //Als PNG-Datei speichern
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                encoder.Save(fs);
            }
        }

        public void speichern(string ExcelFilePath, string Seite)
        {          
            //make sure that the Data Exists
            if (File.Exists(ExcelFilePath))
            {
                try
                {

                    using (var package = new ExcelPackage(ExcelFilePath))
                    {
                        var worksheet = package.Workbook.Worksheets[0];

                        object Dokument = null; // Dokument in dem das Aktuell offene Formular gespeichert wird

                        //Create a Dictionary for Cell Mappings
                        Dictionary<string, string> CellMappings = new Dictionary<string, string>();

                        //Switch Funktion for Selecting the right document and CellMappings
                        switch (Seite)
                        {
                            case "IbnP":
                                Dokument = CC.Content as Inbetriebnahme_Protokoll;
                                CellMappings = GlobalVariables.CellMapping_IbnP;
                                break;

                            case "Serviceanforderungen":
                                Dokument = CC.Content as Service_Anforderung;
                                CellMappings = GlobalVariables.CellMapping_ServiceAnforderungen;
                                break;

                            case "Stundennachweis":
                                Dokument = CC.Content as Stundennachweis;
                                CellMappings = GlobalVariables.CellMapping_Stundenachweis;
                                break;

                            case "Interner_Bericht":
                                Dokument = CC.Content as Interner_Bericht;
                                CellMappings = GlobalVariables.CellMapping_InternerBericht;
                                break;

                            case "IbnP_MRS":
                                Dokument = CC.Content as Inbetriebnahmeprotokoll_MRS;
                                CellMappings = GlobalVariables.CellMapping_IBNP_MRS;
                                break;
                        }

                        foreach (KeyValuePair<string, string> CellMapping in CellMappings) // Schleife über die Länge des ZellenObjekte Arrays
                        {
                            string Zelle = CellMapping.Key;

                            string Objectname = CellMapping.Value;

                            string Object_bezeichnung = Objectname.Substring(0, 2).ToUpper(); // hier werden die ersten zwei Buchstaben der Object namen abgetrennt da man anhand dessen die Objekttypen unterscheiden kann


                            if (Dokument is FrameworkElement element)
                            {
                                // Select the Correct Objecttype based on the and Object_bezeichnung
                                switch (Object_bezeichnung)
                                {
                                    //If the right Objecttype is found search for the Object in the UserControl and save the Value in the Excel Cell
                                    case "TB":
                                        TextBox objekt_tb = element.FindName(Objectname) as TextBox;

                                        worksheet.Cells[Zelle].Value = objekt_tb.Text;

                                        break;

                                    case "CB":
                                        ComboBox objekt_cb = element.FindName(Objectname) as ComboBox;

                                        worksheet.Cells[Zelle].Value = objekt_cb.Text;

                                        break;

                                    case "DP":
                                        DatePicker objekt_dp = element.FindName(Objectname) as DatePicker;

                                        worksheet.Cells[Zelle].Value = objekt_dp.Text;

                                        break;

                                    case "CH":
                                        CheckBox objekt_ch = element.FindName(Objectname) as CheckBox;

                                        if (objekt_ch.IsChecked == true)
                                        {
                                            worksheet.Cells[Zelle].Value = "X";
                                        }
                                        else
                                        {
                                            worksheet.Cells[Zelle].Value = "";
                                        }
                                        break;
                                }
                            }
                        }
                        package.Save();
                        package.Dispose();
                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Die Datei '" + ExcelFilePath + "' kann nicht gespeichert werden. Sie ist möglicherweise noch geöffnet oder schreibgeschützt.\n\n" + ex.Message, "Fehler beim Speichern", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Beim Speichern ist ein Fehler aufgetreten:\n" + ex.Message, "Fehler beim Speichern", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        } // allgemeine Speicherfunktion die auf jedes Dokument anwendbar ist

        public void Laden(string ExcelFilePath, string Seite)
        {
            //Check if the Excel file exists at the specified path
            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Same as in Save Funktion, with the difference that the values of the Excel-Data are loaded into the UserControl
                using (var package = new ExcelPackage(ExcelFilePath))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    object Dokument = null;

                    Dictionary<string, string> CellMappings = new Dictionary<string, string>();

                    switch (Seite)
                    {
                        case "IbnP":
                            Dokument = CC.Content as Inbetriebnahme_Protokoll;
                            CellMappings = GlobalVariables.CellMapping_IbnP;
                            break;

                        case "Serviceanforderungen":
                            Dokument = CC.Content as Service_Anforderung;
                            CellMappings = GlobalVariables.CellMapping_ServiceAnforderungen;
                            break;

                        case "Stundennachweis":
                            Dokument = CC.Content as Stundennachweis;
                            CellMappings = GlobalVariables.CellMapping_Stundenachweis;
                            break;

                        case "Interner_Bericht":
                            Dokument = CC.Content as Interner_Bericht;
                            CellMappings = GlobalVariables.CellMapping_InternerBericht;
                            break;

                        case "IbnP_MRS":
                            Dokument = CC.Content as Inbetriebnahmeprotokoll_MRS;
                            CellMappings = GlobalVariables.CellMapping_IBNP_MRS;
                            break;
                    }

                    foreach (KeyValuePair<string, string> CellMapping in CellMappings) // Schleife über die Länge des ZellenObjekte Arrays
                    {
                        string Zelle = CellMapping.Key;

                        string Objectname = CellMapping.Value;

                        string Object_bezeichnung = Objectname.Substring(0, 2).ToUpper();

                        

                        if (Dokument is FrameworkElement element) 
                        {
                            switch (Object_bezeichnung)
                            {
                                case "TB":
                                    TextBox objekt_tb = element.FindName(Objectname) as TextBox;

                                    objekt_tb.Text = worksheet.Cells[Zelle].Text;

                                    break;

                                case "CB":
                                    ComboBox objekt_cb = element.FindName(Objectname) as ComboBox;

                                    objekt_cb.Text = worksheet.Cells[Zelle].Text;

                                    break;

                                case "DP":
                                    DatePicker objekt_dp = element.FindName(Objectname) as DatePicker;

                                    objekt_dp.Text = worksheet.Cells[Zelle].Text;

                                    break;

                                case "CH":
                                    CheckBox objekt_ch = element.FindName(Objectname) as CheckBox;

                                    if (worksheet.Cells[Zelle].Text == "X")
                                    {
                                        objekt_ch.IsChecked = true;
                                    }
                                    else
                                    {
                                        objekt_ch.IsChecked= false;
                                    }

                                    break;
                            }
                        }
                    }
                    package.Dispose();
                }
            }
        } // allgemeine Ladefunktion die auf jedes Dokument anwendbar ist

        private void Auftrag_Downloaden(object sender, RoutedEventArgs e)
        {
            //Funktion to get all Data of the Current Order and save it in the Local Order Folder. so that the User can work Offline
            string Pfad_DokumentOrdner = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);// Get the path to the user's Documents folder
            string Pfad_Servicetool_Lokal = string.Format(Properties.Resources.Pfad_AuftragsOrdner_Off, GlobalVariables.AuftragsNR);
            string Lokaler_Pfad = System.IO.Path.Combine(Pfad_DokumentOrdner, Pfad_Servicetool_Lokal);

            //Check if the Local Order Folder already exists, if not create it
            if (Directory.Exists(Lokaler_Pfad))
            {
                //copy all files from the Global Order Folder to the Local Order Folder
                foreach (string datei in Directory.GetFiles(GlobalVariables.Pfad_AuftragsOrdner))
                {
                    string DateiName = System.IO.Path.GetFileName(datei);
                    string ZielDateiPfad = System.IO.Path.Combine(Lokaler_Pfad, DateiName);

                    if (!File.Exists(ZielDateiPfad))
                    {
                        File.Copy(datei, ZielDateiPfad);
                    }
                }
                //Create the Subfolder for Attachments and Signatures if it does not exist
                Directory.CreateDirectory(Path.Combine(Lokaler_Pfad,"Anhaenge/Unterschriften"));
            }
            else 
            {
                Directory.CreateDirectory(Lokaler_Pfad);

                foreach (string datei in Directory.GetFiles(GlobalVariables.Pfad_AuftragsOrdner))
                {
                    string DateiName = System.IO.Path.GetFileName(datei);
                    string ZielDateiPfad = System.IO.Path.Combine(Lokaler_Pfad, DateiName);

                    if (!File.Exists(ZielDateiPfad))
                    {
                        File.Copy(datei, ZielDateiPfad);
                    }
                }
                Directory.CreateDirectory(Path.Combine(Lokaler_Pfad, "Anhaenge/Unterschriften"));
            }
        }

        private void CB_Sprache_auswahl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //ToDo need to introduce the Funktion to switch the Language of the Application/Current Window
        }

        private void Creat_PDF_Dokuments(object sender, EventArgs e) 
        {
            MessageBox.Show("Stelle sicher das alle Dokumente vollständig ausgefüllt sind!");
            EngDe_For_PDF Translations = GetLanguageDataForPDF(); //Load all Translations for the PDF Documents
            int LanguageNumber = 0; // Variable to define the Language, 0 = Deutsch, 1 = Englisch
            if (GlobalVariables.Sprache_Kunde == "Deutsch")
            {
                LanguageNumber = 1;
            }
            else
            {
                LanguageNumber = 0;
            }
            //Create all PDF Documents
            Create_PDF_Of_Stundennachweis(Translations, LanguageNumber);
            Create_PDF_Of_Inbetriebnahmeprotokoll(Translations, LanguageNumber);
            Create_PDF_Of_IbnP_MRS(Translations, LanguageNumber);
        }
        
        public EngDe_For_PDF GetLanguageDataForPDF()
        {
            string json = File.ReadAllText(Properties.Resources.Pfad_LanguageForPDF);
        
            List<TranslationEntryPDF> sprachtabelleEntries = JsonConvert.DeserializeObject<List<TranslationEntryPDF>>(json);
            EngDe_For_PDF engDe_For_PDF = new EngDe_For_PDF();

            foreach(TranslationEntryPDF entry in sprachtabelleEntries)
            {
                // Search for Property in EngDe_For_PDF class by name
                PropertyInfo property = typeof(EngDe_For_PDF).GetProperty(entry.VariableName);

                if (property != null && property.PropertyType == typeof(List<string>))
                {
                    
                    var list = (List<string>)property.GetValue(engDe_For_PDF);

                    // Füge die Werte hinzu
                    list.Add(entry.Deutsch);
                    list.Add(entry.Englisch);
                }
                else
                {
                    Console.WriteLine($"Warnung: Property '{entry.VariableName}' nicht gefunden oder falscher Typ!");
                }
            }
            //Return Object of Class for Translations
            return engDe_For_PDF;
        }

        private void Create_PDF_Of_Stundennachweis(EngDe_For_PDF Translations, int LanguageNumber)
        {
            //Save all Information out off the excell data that are needed for the PDF
            Stundennachweis_PDF_Data pDF_Data = GetDataForPDF_StdN();
            QuestPDF.Settings.License = LicenseType.Community;
            string SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.pdf");

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(35);
                    page.Size(PageSizes.A4);
                    page.PageColor(QuestPDF.Helpers.Colors.White);
                    page.Header().PaddingBottom(10).BorderBottom(1).Column(column => //Places a Colum in witch the Items are Vertical aligned
                    {
                        
                        column.Item().Row(row => // adding a row in witch the Items are Horizontal aligned
                        {
                            row.RelativeItem().Text(Translations.ServiceVisitReport[LanguageNumber]).FontSize(27).Bold();
                            row.ConstantItem(100)
                            .AlignRight()
                            .Image("Bilder/gneuss_png_1.png");
                        });

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Column(col=>
                            {
                                col.Item().Text(Translations.OrderNo[LanguageNumber] + GlobalVariables.AuftragsNR).FontSize(12);
                                col.Item().Text(Translations.Customer[LanguageNumber] + pDF_Data.Customer).FontSize(12);
                                col.Item().Text(Translations.ContactPerson[LanguageNumber] + pDF_Data.ContactPerson).FontSize(12);
                            });


                            row.RelativeItem().AlignRight().Column(col =>
                            {
                                col.Item().Text(Translations.Date[LanguageNumber] + DateTime.Now.ToString("dd.MM.yyyy")).FontSize(12);
                                col.Item().Text(Translations.Technican[LanguageNumber] + pDF_Data.ServiceTechnician).FontSize(12);
                                col.Item().Text(Translations.Adress[LanguageNumber] + pDF_Data.Adress1).FontSize(12);
                                col.Item().Text("                 " + pDF_Data.Adress2).FontSize(12);
                            });                            
                        });
                        
                    });

                    page.Footer().PaddingTop(10).BorderTop(1).Row(row =>
                    {
                        row.RelativeItem().Text("Gneuss Kunststofftechnik GmbH - Moenichhusen 42 - 32549 Bad Oeynhausen - Germany \n                           Phone:+49 57 31/5 30 70 - Fax:+49 57 31/53 07-77").FontSize(9).SemiBold();
                        row.ConstantItem(100).AlignRight().Text(text =>
                        {
                            
                            text.Span("Seite ").FontSize(9);
                            text.CurrentPageNumber().FontSize(9); // Aktuelle Seitennummer
                            text.Span(" von ").FontSize(9);
                            text.TotalPages().FontSize(9); // Gesamtanzahl der Seiten
                        });
                        
                    });

                    page.Content()
                    .Column(column =>
                    {
                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.TravelTime[LanguageNumber]).FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.MeansOfTransport[LanguageNumber] + pDF_Data.meansofTransport).FontSize(12).AlignCenter();
                        });

                        column.Spacing(5);

                        Func<IContainer, IContainer> headerstyle = c => c
                            .Background(QuestPDF.Helpers.Colors.Grey.Lighten2)
                            .BorderBottom(1).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1)
                            .PaddingVertical(4).PaddingHorizontal(2)
                            .AlignCenter().AlignMiddle();

                        Func<IContainer, IContainer> cellstyle = c => c
                            .Border(0.5f).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken2)
                            .AlignCenter().AlignMiddle();

                        column.Item().Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(60);
                                columns.ConstantColumn(70);
                                columns.ConstantColumn(50);
                                columns.ConstantColumn(70);
                                columns.ConstantColumn(50);
                                columns.ConstantColumn(50);
                                columns.ConstantColumn(60);
                                columns.ConstantColumn(70);
                            });

                            
                            //Border(1).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1)
                            table.Header(header =>
                            {
                                header.Cell().Text("   ").FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.DateStart[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.TimeStart[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.DateEnd[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.TimeEnd[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.Break[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text(Translations.TotalTime[LanguageNumber]).FontSize(14);
                                header.Cell().Element(headerstyle).Text("Kilometer").FontSize(14);
                            });

                            

                            table.Cell().Element(cellstyle).Text(Translations.Arrival[LanguageNumber]).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateArrivalStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeArrivalStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateArrivalEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeArrivalEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.BreakArrival).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalTimeArrival).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalKilometersArrival).FontSize(12);

                            table.Cell().Element(cellstyle).Text(Translations.Departure[LanguageNumber]).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateDepartureStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeDepartureStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateDepartureEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeDepartureEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.BreakDeparture).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalTimeDeparture).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalKilometersDeparture).FontSize(12);
                                                       
                        });
                        column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Spacing(5);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.WorkingTime[LanguageNumber]).FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);

                        column.Item().Table(table =>
                        {
                            table.ColumnsDefinition(columns =>
                            {
                                columns.ConstantColumn(60);
                                columns.ConstantColumn(48);
                                columns.ConstantColumn(48);
                                columns.ConstantColumn(45);
                                columns.ConstantColumn(70);
                                columns.ConstantColumn(60);
                                columns.ConstantColumn(60);
                                columns.ConstantColumn(69);
                                columns.ConstantColumn(50);
                            });

                            table.Header(header =>
                            {
                                header.Cell().Element(headerstyle).Text(Translations.Date[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.TimeStart[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.TimeEnd[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.Break[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.Note[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.Normalhours[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.OverTime[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.Nightwork[LanguageNumber]).FontSize(12);
                                header.Cell().Element(headerstyle).Text(Translations.TotalTime[LanguageNumber]).FontSize(12);
                            });

                            foreach(StundenTabellenEintrag workday in pDF_Data.Arbeitszeit)
                            {
                                table.Cell().Element(cellstyle).Text(workday.Date).FontSize(10);
                                table.Cell().Element(cellstyle).Text(workday.Start).FontSize(10);
                                table.Cell().Element(cellstyle).Text(workday.End).FontSize(10);
                                table.Cell().Element(cellstyle).Text(workday.Break).FontSize(10);
                                

                                if (workday.StartS2 != "" && workday.EndS2 != "")
                                {
                                    table.Cell().RowSpan(2).Element(cellstyle).Text(workday.Note).FontSize(10);
                                    table.Cell().RowSpan(2).Element(cellstyle).Text(workday.NormalStunden).FontSize(10);
                                    table.Cell().RowSpan(2).Element(cellstyle).Text(workday.OverTime).FontSize(10);
                                    table.Cell().RowSpan(2).Element(cellstyle).Text(workday.Nightwork).FontSize(10);
                                    table.Cell().RowSpan(2).Element(cellstyle).Text(workday.TotalHours).FontSize(10);
                                    table.Cell().Element(cellstyle).Text("Schicht 2").FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.StartS2).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.EndS2).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.BreakS2).FontSize(10);
                                }
                                else
                                {
                                    table.Cell().Element(cellstyle).Text(workday.Note).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.NormalStunden).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.OverTime).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.Nightwork).FontSize(10);
                                    table.Cell().Element(cellstyle).Text(workday.TotalHours).FontSize(10);
                                }

                                
                            }
                            table.Cell().ColumnSpan(5).Border(0.5f).BorderColor(QuestPDF.Helpers.Colors.Black).Background(QuestPDF.Helpers.Colors.Grey.Darken1).AlignRight().Text("Summe ");
                            table.Cell().Border(1).BorderColor(QuestPDF.Helpers.Colors.Black).Text(" " + FormattedTimeSpanInHHMM(pDF_Data.TotalNormalHours));
                            table.Cell().Border(1).BorderColor(QuestPDF.Helpers.Colors.Black).Text(" " + FormattedTimeSpanInHHMM(pDF_Data.TotalOverTime));
                            table.Cell().Border(1).BorderColor(QuestPDF.Helpers.Colors.Black).Text(" " + FormattedTimeSpanInHHMM(pDF_Data.TotalNightWork));
                            table.Cell().Border(1).BorderColor(QuestPDF.Helpers.Colors.Black).Text(" " + FormattedTimeSpanInHHMM(pDF_Data.TotalHours));
                        });

                        column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Spacing(15);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.Report[LanguageNumber]).FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);
                        column.Item().Row(row =>
                        {
                            int countReports = pDF_Data.Report.Count;
                            string allReports = "";
                            for (int i = 0; i < countReports; i++)
                            {
                                if(LanguageNumber == 0) // Deutsch
                                    allReports += $"Woche {i+1}:\n" + pDF_Data.Report[i] + "\n";
                                else // Englisch
                                    allReports += $"Week " + (i+1) + ":\n" + pDF_Data.Report[i] + "\n";
                            }
                            row.RelativeItem().Text(allReports);
                        });

                        column.Item().PageBreak();

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.Training[LanguageNumber]).FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);


                        column.Item().Row(row => {
                            row.AutoItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.SetupAufbau + Translations.Setup[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.OperatingPrinciple + Translations.Operatingprinciple[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.BriefingControlSystem + Translations.Acceptance[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.Sonstiges + Translations.Training[LanguageNumber]).FontSize(10).AlignLeft();
                            });

                            row.Spacing(50);

                            row.RelativeItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.Troubleshooting + Translations.Troubleshooting[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.OperationOfWholeEquipment + Translations.OperatingOfWholeEquipment[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.SafetyInstructions + Translations.SafetyInstructions[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.Sonstiges + Translations.Maintenance[LanguageNumber]).FontSize(10).AlignLeft();
                            });
                        });

                        column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.Evaluation[LanguageNumber]).FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);
                        //"😊", "😐", "🙁"

                        column.Item().Row(row =>
                        {
                            row.AutoItem().Column(col =>
                            { 
                                col.Item().Text(Translations.HowSatisfiedWithTheService[LanguageNumber]).FontSize(10).AlignLeft();
                                col.Item().Text(" ").FontSize(10).AlignLeft();
                                col.Item().Text(Translations.HowSatisfiedWithTheSupport[LanguageNumber]).FontSize(10).AlignLeft();
                            });

                            row.Spacing(15);

                            row.AutoItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.EvaluationProduktGood).FontSize(16).AlignLeft();
                                col.Item().Text("😊").FontSize(16).AlignLeft();
                                col.Item().Text(pDF_Data.EvaluationSupportGood).FontSize(16).AlignLeft();
                            });

                            row.Spacing(15);

                            row.AutoItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.EvaluationProduktMid).FontSize(16).AlignLeft();
                                col.Item().Text("😐").FontSize(16).AlignLeft();
                                col.Item().Text(pDF_Data.EvaluationSupportMid).FontSize(16).AlignLeft();
                            });

                            row.Spacing(15);

                            row.AutoItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.EvaluationProduktBad).FontSize(16).AlignLeft();
                                col.Item().Text("🙁").FontSize(16).AlignLeft();
                                col.Item().Text(pDF_Data.EvaluationSupportBad).FontSize(16).AlignLeft();
                            });
                        });
                        column.Spacing(10);

                        column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Item().Text(Translations.Signatures[LanguageNumber]).FontSize(20).AlignCenter();

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(Translations.MachineAcceptanceaccordingToOrderConfirmation[LanguageNumber]).FontSize(10).AlignCenter();
                        });

                        column.Item().Row(row =>
                        {
                            row.AutoItem().Border(1).Column(col =>
                            {
                                col.Spacing(0);
                                col.Item().Width(135).Height(25).Border(1).AlignMiddle().Text(Translations.Place[LanguageNumber]).FontSize(12).AlignCenter();
                                col.Item().Width(135).Height(40).Border(1).Text(pDF_Data.PlaceCustomerSignature).FontSize(10).AlignCenter();                                
                            });
                            row.AutoItem().Column(col =>
                            {
                                col.Item().Width(115).Height(25).Border(1).AlignMiddle().Text(Translations.PlaceDate[LanguageNumber]).FontSize(12).AlignCenter();
                                col.Item().Width(115).Height(40).Border(1).Text(pDF_Data.Date_Technican_Signature).FontSize(10).AlignCenter();

                            });
                            row.AutoItem().Column(col =>
                            {
                                string imagepath_sign_technican = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureemployee.png");
                                col.Item().Width(265).Height(25).Border(1).AlignMiddle().Text(Translations.Technican[LanguageNumber]).FontSize(12).AlignCenter();
                                if(File.Exists(imagepath_sign_technican))col.Item().Width(265).Height(40).Border(1).AlignRight().Image(imagepath_sign_technican);
                            });
                        });
                        column.Item().Row(row =>
                        {
                            row.AutoItem().Column(col =>
                            {                                
                                col.Spacing(0);
                                col.Item().Width(135).Height(25).Border(1).AlignMiddle().Text(Translations.Place[LanguageNumber]).FontSize(12).AlignCenter();
                                col.Item().Width(135).Height(40).Border(1).Text(pDF_Data.PlaceCustomerSignature).FontSize(10).AlignCenter();

                            });
                            
                            row.AutoItem().Column(col =>
                            {
                                col.Item().Width(115).Height(25).Border(1).AlignMiddle().Text(Translations.Date[LanguageNumber]).FontSize(12).AlignCenter();
                                col.Item().Width(115).Height(40).Border(1).Text(pDF_Data.Date_Customer_Signature).FontSize(10).AlignCenter();
                            });
                            row.AutoItem().Column(col =>
                            {
                                string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "StdNSignatureCustomer.png");
                                col.Item().Width(265).Height(25).Border(1).AlignMiddle().Text(Translations.Customer[LanguageNumber]).FontSize(12).AlignCenter();
                                if (File.Exists(ImagePath_Sign_Kunde)) col.Item().Width(265).Height(40).Border(1).AlignRight().Image(ImagePath_Sign_Kunde);
                            });

                        });
                    });
                });
            });
            document.GeneratePdf(SavePath);

        }
        public Stundennachweis_PDF_Data GetDataForPDF_StdN()
        {
            //Calculate the Service Duration in Days
            TimeSpan ServiceDurationInDays = GlobalVariables.EndeServiceEinsatz - GlobalVariables.StartServiceEinsatz;
            //Convert the Service Duration in Days to Weeks
            double weeksnotRounded = ServiceDurationInDays.TotalDays / 7;
            int NumberOfStundennachweis = (int)Math.Ceiling(weeksnotRounded);

            TimeSpan TotalNormalStd = new TimeSpan();
            TimeSpan TotalOverTime = new TimeSpan();
            TimeSpan TotalNightwork = new TimeSpan();
            TimeSpan TotalTime = new TimeSpan();


            Stundennachweis_PDF_Data PDF_Data = new Stundennachweis_PDF_Data();
            for (int i = 0; i < NumberOfStundennachweis; i++) // Loop through the number of Stundennachweis documents
            {
                string ExcelFilePath = "";
                
                if (i == 0) // if its is the first document use Stundennachweis.xlsm else count up from StdN_2
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm");
                    //implement the License Context for EPPlus
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; 

                        //Safe Information for the Header of the PDF
                        PDF_Data.Customer = worksheet.Cells["C4"].Text;
                        PDF_Data.ServiceTechnician = worksheet.Cells["P1"].Text;
                        PDF_Data.Adress1 = worksheet.Cells["M4"].Text;
                        PDF_Data.Adress2 = worksheet.Cells["M5"].Text;
                        PDF_Data.ContactPerson = worksheet.Cells["J6"].Text;
                        PDF_Data.meansofTransport = worksheet.Cells["E8"].Text;
                        PDF_Data.DateArrivalStart = worksheet.Cells["C10"].Text;
                        PDF_Data.DateArrivalEnd = worksheet.Cells["H10"].Text;
                        PDF_Data.TimeArrivalStart = worksheet.Cells["E10"].Text;
                        PDF_Data.TimeArrivalEnd = worksheet.Cells["J10"].Text;
                        PDF_Data.BreakArrival = worksheet.Cells["L10"].Text;
                        PDF_Data.TotalTimeArrival = worksheet.Cells["O10"].Text;
                        PDF_Data.TotalKilometersArrival = worksheet.Cells["Q10"].Text;
                        PDF_Data.DateDepartureStart = worksheet.Cells["C12"].Text;
                        PDF_Data.DateDepartureEnd = worksheet.Cells["H12"].Text;
                        PDF_Data.TimeDepartureStart = worksheet.Cells["E12"].Text;
                        PDF_Data.TimeDepartureEnd = worksheet.Cells["J12"].Text;
                        PDF_Data.BreakDeparture = worksheet.Cells["L12"].Text;
                        PDF_Data.TotalTimeDeparture = worksheet.Cells["O12"].Text;
                        PDF_Data.TotalKilometersDeparture = worksheet.Cells["Q12"].Text;
                        PDF_Data.Report.Add(worksheet.Cells["A35"].Text);
                        if (worksheet.Cells["A65"].Text.ToUpper() == "X") { PDF_Data.SetupAufbau = "☑"; } else { PDF_Data.SetupAufbau = "☐"; }
                        if (worksheet.Cells["A66"].Text.ToUpper() == "X") { PDF_Data.OperatingPrinciple = "☑"; } else { PDF_Data.OperatingPrinciple = "☐"; }
                        if (worksheet.Cells["A67"].Text.ToUpper() == "X") { PDF_Data.BriefingControlSystem = "☑"; } else { PDF_Data.BriefingControlSystem = "☐"; }
                        if (worksheet.Cells["A68"].Text.ToUpper() == "X") { PDF_Data.Sonstiges = "☑"; } else { PDF_Data.Sonstiges = "☐"; }
                        if (worksheet.Cells["G64"].Text.ToUpper() == "X") { PDF_Data.Troubleshooting = "☑"; } else { PDF_Data.Troubleshooting = "☐"; }
                        if (worksheet.Cells["G65"].Text.ToUpper() == "X") { PDF_Data.OperationOfWholeEquipment = "☑"; } else { PDF_Data.OperationOfWholeEquipment = "☐"; }
                        if (worksheet.Cells["G66"].Text.ToUpper() == "X") { PDF_Data.SafetyInstructions = "☑"; } else { PDF_Data.SafetyInstructions = "☐"; }
                        if (worksheet.Cells["G67"].Text.ToUpper() == "X") { PDF_Data.Maintenance = "☑"; } else { PDF_Data.Maintenance = "☐"; }

                        if (worksheet.Cells["K72"].Text.ToUpper() == "X") { PDF_Data.EvaluationProduktGood = "☑"; } else { PDF_Data.EvaluationProduktGood = "☐"; }
                        if (worksheet.Cells["M72"].Text.ToUpper() == "X") { PDF_Data.EvaluationProduktMid = "☑"; } else { PDF_Data.EvaluationProduktMid = "☐"; }
                        if (worksheet.Cells["O72"].Text.ToUpper() == "X") { PDF_Data.EvaluationProduktBad = "☑"; } else { PDF_Data.EvaluationProduktBad = "☐"; }
                        if (worksheet.Cells["K74"].Text.ToUpper() == "X") { PDF_Data.EvaluationSupportGood = "☑"; } else { PDF_Data.EvaluationSupportGood = "☐"; }
                        if (worksheet.Cells["M74"].Text.ToUpper() == "X") { PDF_Data.EvaluationSupportMid = "☑"; } else { PDF_Data.EvaluationSupportMid = "☐"; }
                        if (worksheet.Cells["O74"].Text.ToUpper() == "X") { PDF_Data.EvaluationSupportBad = "☑"; } else { PDF_Data.EvaluationSupportBad = "☐"; }

                        PDF_Data.Date_Technican_Signature = worksheet.Cells["G70"].Text;
                        PDF_Data.Date_Customer_Signature = worksheet.Cells["G77"].Text;
                        PDF_Data.PlaceCustomerSignature = worksheet.Cells["C76"].Text;
                        //☑☐
                        if (worksheet.Cells["J31"].Text != "") 
                        { 
                            TotalNormalStd += TimeSpan.Parse(worksheet.Cells["J31"].Text);
                            TotalOverTime += TimeSpan.Parse(worksheet.Cells["M31"].Text);
                            TotalNightwork += TimeSpan.Parse(worksheet.Cells["O31"].Text);
                            TotalTime += TimeSpan.Parse(worksheet.Cells["Q31"].Text);
                        }
                        for (int x = 0; x < 14; x+=2)
                        {
                            if(worksheet.Cells["C" + (x + 17)].Text != "")
                            {
                                PDF_Data.Arbeitszeit.Add(new StundenTabellenEintrag
                                {
                                    Date = worksheet.Cells["B" + (x + 17)].Text,
                                    Start = worksheet.Cells["C" + (x + 17)].Text,
                                    End = worksheet.Cells["D" + (x + 17)].Text,
                                    Break = worksheet.Cells["E" + (x + 17)].Text,
                                    StartS2 = worksheet.Cells["C" + (x + 18)].Text,
                                    EndS2 = worksheet.Cells["D" + (x + 18)].Text,
                                    BreakS2 = worksheet.Cells["E" + (x + 18)].Text,
                                    Note = worksheet.Cells["F" + (x + 17)].Text,
                                    NormalStunden = worksheet.Cells["J" + (x + 17)].Text,
                                    OverTime = worksheet.Cells["M" + (x + 17)].Text,
                                    Nightwork = worksheet.Cells["O" + (x + 17)].Text,
                                    TotalHours = worksheet.Cells["Q" + (x + 17)].Text,
                                });
                            }
                            
                        }

                    }
                }
                else
                {// If its not the first document only add the Workingtime to the Arbeitszeit List Object of the Class
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_" + (i + 1) + ".xlsm");
                    using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; 
                        PDF_Data.Report.Add(worksheet.Cells["A35"].Text);
                        if (worksheet.Cells["J31"].Text != "")
                        {
                            TotalNormalStd += TimeSpan.Parse(worksheet.Cells["J31"].Text);
                            TotalOverTime += TimeSpan.Parse(worksheet.Cells["M31"].Text);
                            TotalNightwork += TimeSpan.Parse(worksheet.Cells["O31"].Text);
                            TotalTime += TimeSpan.Parse(worksheet.Cells["Q31"].Text);
                        }
                        
                        for (int x = 0; x < 14; x+=2)
                        {
                            if (worksheet.Cells["C" + (x + 17)].Text != "")
                            {
                                PDF_Data.Arbeitszeit.Add(new StundenTabellenEintrag
                                {
                                    Date = worksheet.Cells["B" + (x + 17)].Text,
                                    Start = worksheet.Cells["C" + (x + 17)].Text,
                                    End = worksheet.Cells["D" + (x + 17)].Text,
                                    Break = worksheet.Cells["E" + (x + 17)].Text,
                                    StartS2 = worksheet.Cells["C" + (x + 18)].Text,
                                    EndS2 = worksheet.Cells["D" + (x + 18)].Text,
                                    BreakS2 = worksheet.Cells["E" + (x + 18)].Text,
                                    Note = worksheet.Cells["F" + (x + 17)].Text,
                                    NormalStunden = worksheet.Cells["J" + (x + 17)].Text,
                                    OverTime = worksheet.Cells["M" + (x + 17)].Text,
                                    Nightwork = worksheet.Cells["O" + (x + 17)].Text,
                                    TotalHours = worksheet.Cells["Q" + (x + 17)].Text,
                                });
                            }
                        }
                    }
                }
            }
            PDF_Data.TotalNormalHours = TotalNormalStd;
            PDF_Data.TotalOverTime = TotalOverTime;
            PDF_Data.TotalNightWork = TotalNightwork;
            PDF_Data.TotalHours = TotalTime;
            return PDF_Data;//return the PDF Data Object with all the information witch are inserted in the Class
        }

        public void Create_PDF_Of_Inbetriebnahmeprotokoll(EngDe_For_PDF Translations, int LanguageNumber)
        {
            int NumberOfIbnP = 0;
            string ExcelFilePath = "";
            string SavePath = "";
            if (GlobalVariables.Maschiene_1 != "" && GlobalVariables.Maschiene_1 != "MRS" && GlobalVariables.Maschiene_1 != "Jump" && GlobalVariables.Maschiene_1 != "3C-RF")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_2 != "" && GlobalVariables.Maschiene_2 != "MRS" && GlobalVariables.Maschiene_2 != "Jump" && GlobalVariables.Maschiene_2 != "3C-RF")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_3 != "" && GlobalVariables.Maschiene_3 != "MRS" && GlobalVariables.Maschiene_3 != "Jump" && GlobalVariables.Maschiene_3 != "3C-RF")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_4 != "" && GlobalVariables.Maschiene_4 != "MRS" && GlobalVariables.Maschiene_4 != "Jump" && GlobalVariables.Maschiene_4 != "3C-RF")
            {
                NumberOfIbnP++;
            }
            for (int i = 0; i < NumberOfIbnP; i++)
            {
                
                if (i == 0)
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                    SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.pdf");
                }
                else
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_" + (i + 1) + ".xlsm");
                    SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_" + (i + 1) + ".pdf");
                }
                PDF_Data_InbetriebnahmeProtokoll pDF_Data = GetDataForIbnP_PDF(ExcelFilePath);
                QuestPDF.Settings.License = LicenseType.Community;
                
                var Dokument = Document.Create(document =>
                {
                    Func<IContainer, IContainer> headerstyle = c => c
                        .Background(QuestPDF.Helpers.Colors.Grey.Lighten2)
                        .Border(1).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1)
                        .PaddingVertical(2).AlignCenter().AlignMiddle();

                    Func<IContainer, IContainer> cellstyle = c => c
                        .Border(0.5f).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken2)
                        .AlignCenter().AlignMiddle();

                    document.Page(page =>
                    {
                        page.Size(PageSizes.A4.Landscape());
                        page.Margin(10);
                        page.PageColor(QuestPDF.Helpers.Colors.White);
                        page.Header().PaddingBottom(5).BorderBottom(1).Column(column =>
                        {
                            column.Spacing(5);
                            column.Item().Row(row =>
                            {
                                row.RelativeItem().Text(Translations.CommissioningDataSheet[LanguageNumber]).FontSize(20).SemiBold().AlignCenter();
                                row.ConstantItem(100)
                                .AlignRight()
                                .Image("Bilder/gneuss_png_1.png");
                            });
                            column.Spacing(15);
                            
                        });
                        page.Footer().PaddingTop(10).BorderTop(1).Row(row =>
                        {
                            row.RelativeItem().Text("Gneuss Kunststofftechnik GmbH - Moenichhusen 42 - 32549 Bad Oeynhausen - Germany Phone:+49 57 31/5 30 70 - Fax:+49 57 31/53 07-77").FontSize(9).SemiBold();
                            row.ConstantItem(100).AlignRight().Text(text =>
                            {
                                text.Span("Seite ").FontSize(9);
                                text.CurrentPageNumber().FontSize(9); // Aktuelle Seitennummer
                                text.Span(" von ").FontSize(9);
                                text.TotalPages().FontSize(9); // Gesamtanzahl der Seiten
                            });

                        });
                        page.Content().PaddingVertical(5).Column(column =>
                        {
                            column.Item().AlignCenter().Text("Auftragsinformationen / Orderinformation").FontSize(16).Underline();
                            
                            column.Item().Row(row =>
                            {
                                row.RelativeItem().Column(col =>
                                {
                                    col.Item().Text(Translations.Customer[LanguageNumber] + pDF_Data.Customer).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.ContactPerson[LanguageNumber] + pDF_Data.ContactPerson).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.OrderNo[LanguageNumber] + GlobalVariables.AuftragsNR).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.SerialNo[LanguageNumber] + pDF_Data.SerialNumber).FontSize(12).AlignLeft();
                                });
                                row.RelativeItem().Column(col =>
                                {
                                    col.Item().Text(Translations.Filtertype[LanguageNumber] + pDF_Data.FilterType).FontSize(12).AlignLeft();
                                    col.Item().Text("Material: " + pDF_Data.Material).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.Viscosity[LanguageNumber] + pDF_Data.Viscosity).FontSize(12).AlignLeft();
                                    column.Item().Text(Translations.Lineconfiguration[LanguageNumber]).FontSize(12).AlignRight();
                                });
                                row.RelativeItem().Column(col =>
                                {
                                    col.Item().Text(Translations.PreLoading[LanguageNumber] + pDF_Data.Preloading).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.ShimPacking[LanguageNumber] + "LR: " + pDF_Data.ShimpackingLR).FontSize(12).AlignLeft();
                                    col.Item().Text(Translations.ShimPacking[LanguageNumber] + "ZP: " + pDF_Data.ShimpackingZP).FontSize(12).AlignLeft();
                                    column.Item().Text(pDF_Data.LineConfiguration).FontSize(12).AlignLeft();
                                });
                                
                            });

                            column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                            column.Item().AlignCenter().Text(Translations.Processingparameters[LanguageNumber]).FontSize(16).Underline();

                            column.Item().PreventPageBreak().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(25);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                });
                                table.Header(header =>
                                {
                                    header.Cell().Element(headerstyle).Text("lfd.\nNr.\n").FontSize(8);
                                    header.Cell().Element(headerstyle).Text( Translations.PressureP1[LanguageNumber] + "\n" + Translations.UpstreamVF[LanguageNumber] + "\n(bar)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.PressureP1[LanguageNumber] + "\n" + Translations.Downstream[LanguageNumber]+ "\n(bar)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Δp\n \n(bar)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Melt[LanguageNumber] + "\n"+Translations.Temperature[LanguageNumber] + "\ntm(°C)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("n\nExtruder\n(1/min)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Pump[LanguageNumber] + "\n \n(1/min)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Q\n \n(kg/h)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Elements[LanguageNumber] + "\n \n(µm)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.BackFlush[LanguageNumber] + "\n" + Translations.loss[LanguageNumber] + "\n1x(gr)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.BackFlush[LanguageNumber] + "\n" + Translations.loss[LanguageNumber] + "\n10x(gr)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.RSBackflush[LanguageNumber] + "\n" + Translations.LossInPercent[LanguageNumber] + "\n" + Translations.PercentOfQ[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Stroke[LanguageNumber] + "\n" + Translations.Length[LanguageNumber] + "\n(mm)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Back_flush[LanguageNumber] + "\n" + Translations.pressure[LanguageNumber] + "\n(bar)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Drive[LanguageNumber] + "\n" + Translations.force[LanguageNumber] + "\n(bar/kN)").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Flooding[LanguageNumber] + "\n" + Translations.Pin[LanguageNumber] + "\n(mm)").FontSize(8);
                                });
                                for (int x = 0; x < 4; x++)
                                {
                                    table.Cell().Element(cellstyle).Text((x + 1).ToString()).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pressure_P1[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pressure_P2[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.P[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MassTemperatur[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.n_Extruder[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pump[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Q[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.FilterElements[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.BackFlushLoss_1gr[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.BackFlushLoss_10gr[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.BackFlushLossInPercent[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.StrokeLength[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.BackFlushPressure[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.DriveForce[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.FloodingPin[x]).FontSize(8);
                                }
                            });

                            column.Item().PaddingVertical(3).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                            column.Item().AlignCenter().Text(Translations.ScreenchangerControl[LanguageNumber]).FontSize(16).Underline();
                            column.Spacing(5);

                            column.Item().PreventPageBreak().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {

                                    columns.ConstantColumn(84);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                });

                                table.Header(header =>
                                {
                                    header.Cell().Element(headerstyle).Text("RSF Normal\n*Fast Betrieb").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.WStroke[LanguageNumber] + "\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.RStroke[LanguageNumber] + "\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Cycle[LanguageNumber] + "\n ").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.WStroke2[LanguageNumber] + "\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.RStroke2[LanguageNumber] + "\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.PPiston[LanguageNumber] + "\n" + Translations.Forward[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.PPiston[LanguageNumber] + "\n" + Translations.Backwards[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.PPiston2[LanguageNumber] + "\n" + Translations.Forward[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.PPiston2[LanguageNumber] + "\n" + Translations.Backwards[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Number[LanguageNumber] + "\n" + Translations.FilterEle[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).Text(Translations.Strokes[LanguageNumber] + "\n" + Translations.Revolt[LanguageNumber]).FontSize(8);
                                    header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2)
                                        .BorderLeft(1).BorderBottom(1).BorderTop(1).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1)
                                        .PaddingVertical(2).Text("Schuss\nvor    ").AlignRight().FontSize(8);
                                    header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2)
                                        .BorderRight(1).BorderBottom(1).BorderTop(1).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1)
                                        .PaddingVertical(2).Text("kolben %\n    zurück").AlignLeft().FontSize(8);
                                });
                                table.Cell().Element(headerstyle).Text("1").FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke_Filter_RSF_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke_Filter_RSF_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.CycleTime_RSF_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke2_Filter_RSF_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke2_Filter_RSF_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Forward_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Backward_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Forward_2_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Backward_2_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.NumberFilterElements_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.StrokesRevolt_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PuringPiston_Forward_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PuringPiston_Backward_1).FontSize(8);
                                table.Cell().Element(headerstyle).Text("2").FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke_Filter_RSF_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke_Filter_RSF_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.CycleTime_RSF_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke2_Filter_RSF_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke2_Filter_RSF_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Forward_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Backward_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Forward_2_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PPiston_Backward_2_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.NumberFilterElements_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.StrokesRevolt_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PuringPiston_Forward_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PuringPiston_Backward_2).FontSize(8);
                            });

                            column.Item().PreventPageBreak().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(84);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(57);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                    columns.ConstantColumn(56);
                                });
                                table.Header(header =>
                                {
                                    header.Cell().Element(headerstyle).Text("SFX /\nSFXR / SF").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("A.-Hub\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("L.-Hub\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Takt-\nzeit").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Vorflut\nzeit").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Vorflut\nChange").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Vorflut.\nMaß A").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Filter\nSoll-Druck").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Filter\nMin.Druck").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Betriebs-\nart").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Vor/Diff\ndruck").FontSize(8);                                    
                                    header.Cell().Element(headerstyle).Text("Stellung\nRampe").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("Ø Ablauf -\nDüse(mm)").FontSize(8);


                                });
                                table.Cell().Element(headerstyle).Text("1").FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke_Filter_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke_Filter_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.CycleTime_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_Change_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.SetPressure_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.Min_Pressure_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.ModeOfOperation_SFX_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PreDiff_Pressure_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.Flooding_dim_A_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PistonCrossSection_1).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.MeltDischarge_1).FontSize(8);
                                table.Cell().Element(headerstyle).Text("2").FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.WStroke_Filter_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.RStroke_Filter_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.CycleTime_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_Change_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.SetPressure_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.Min_Pressure_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.ModeOfOperation_SFX_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PreDiff_Pressure_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.Flooding_dim_A_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.PistonCrossSection_2).FontSize(8);
                                table.Cell().Element(cellstyle).Text(pDF_Data.MeltDischarge_2).FontSize(8);
                            });

                            column.Item().PreventPageBreak().Row(row =>
                            {
                                row.AutoItem().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(84);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                        columns.ConstantColumn(57);
                                    });
                                    table.Header(header =>
                                    {
                                        header.Cell().Element(headerstyle).Text("KSF").FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.WStroke[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.RStroke[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.Cycle[LanguageNumber] + "\n" + Translations.Time[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.Flooding3[LanguageNumber] + "\n" + Translations.Time[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.PBetween[LanguageNumber] + "\n" + Translations.brPlates[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.Set[LanguageNumber] + "\n" + Translations.Pressure[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.Min[LanguageNumber] + "\n" +Translations.PressureMin[LanguageNumber]).FontSize(8);
                                        header.Cell().Element(headerstyle).Text(Translations.ModeOf + "\n" + Translations.Operation[LanguageNumber]).FontSize(8);
                                    });
                                    table.Cell().Element(headerstyle).Text("1").FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MV_A_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MV_B_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.ScreenLifeTime_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_KSF_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pbetween_br_Plates_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Set_Pressure_KSF_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Min_Pressure_KSF_1).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Mode_Of_Operation_1).FontSize(8);
                                    table.Cell().Element(headerstyle).Text("2").FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MV_A_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MV_B_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.ScreenLifeTime_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.FloodingTime_KSF_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pbetween_br_Plates_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Set_Pressure_KSF_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Min_Pressure_KSF_2).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Mode_Of_Operation_2).FontSize(8);
                                });

                                row.Spacing(5);

                                row.AutoItem().PaddingLeft(5).Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(45);
                                        columns.ConstantColumn(50);
                                        columns.ConstantColumn(40);
                                        columns.ConstantColumn(120);
                                    });
                                    table.Header(header =>
                                    {
                                        header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2).BorderBottom(1).BorderTop(1).BorderLeft(1).Text("VI").FontSize(10).AlignRight();
                                        header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2).BorderBottom(1).BorderTop(1).BorderRight(1).Text("S").FontSize(10).AlignLeft();
                                        header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2).BorderBottom(1).BorderTop(1).BorderLeft(1).Text("Korrekte ").FontSize(10).AlignCenter();
                                        header.Cell().Background(QuestPDF.Helpers.Colors.Grey.Lighten2).BorderBottom(1).BorderRight(1).BorderTop(1).AlignLeft().Text("Funktion der Steuerung").FontSize(10).AlignCenter();
                                    });
                                    table.Cell().Element(headerstyle).Text("VIS").FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.VIS).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Disc_Rotation).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(Translations.discRotation[LanguageNumber]).FontSize(10).AlignCenter();
                                    table.Cell().Element(headerstyle).Text(Translations.dSheet[LanguageNumber]).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.dSheet).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.DriveLoadMeasurement).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(Translations.DriveLoadMeasurement[LanguageNumber]).FontSize(10).AlignCenter();
                                    table.Cell().Element(headerstyle).Text("kp").FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.KP).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.BackflushStrokeLength).FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(Translations.BackFlushStrokeLength[LanguageNumber]).FontSize(10).AlignCenter();
                                    table.Cell().Element(headerstyle).Text("kk").FontSize(10).AlignCenter();
                                    table.Cell().Element(cellstyle).Text(pDF_Data.KK).FontSize(10).AlignCenter();
                                });
                            });

                            column.Item().AlignLeft()
                            .Text(Translations.MaxValueDocumentedAsPhoto + pDF_Data.PhotoAttachment_Yes + Translations.Yes[LanguageNumber] + pDF_Data.PhotoAttachment_No + Translations.NoBecause[LanguageNumber] + pDF_Data.PhotoAttachment_No_Because)
                            .FontSize(14);

                            column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                            column.Item().AlignCenter().Text(Translations.TempProfilInExtrusionsDirection[LanguageNumber]).FontSize(16).Underline();

                            column.Spacing(5);

                            column.Item().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.ConstantColumn(70);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                    columns.ConstantColumn(53);
                                });

                                table.Header(header =>
                                {
                                    header.Cell().Element(headerstyle).Text("Zone").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("1").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("2").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("3").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("4").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("5").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("6").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("7").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("8").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("9").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("10").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("11").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("12").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("13").FontSize(8);
                                    header.Cell().Element(headerstyle).Text("14").FontSize(8);
                                });

                                table.Cell().Element(headerstyle).Text(Translations.Designation[LanguageNumber]).FontSize(8);
                                for (int X = 0; X < 14; X++)
                                {
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Designation_Tempprofil[X]).FontSize(8);
                                }
                                table.Cell().Element(headerstyle).Text("Temperature (°C)").FontSize(8);
                                for (int Y = 0; Y < 14; Y++)
                                {
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Temperatur_Tempprofil[Y]).FontSize(8);
                                }
                            });

                            column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                            column.Item().Row(row => 
                            {
                                row.AutoItem().AlignLeft().Text(pDF_Data.Customer_Temperature_Meassurement_korrekt + Translations.KundenTempMessung[LanguageNumber]).FontSize(12);
                                row.Spacing(10);
                                row.AutoItem().Text(pDF_Data.PressureCutoff + Translations.PressureCutoff[LanguageNumber]).FontSize(12);
                                row.Spacing(10);
                                row.AutoItem().Text(pDF_Data.SetTo + Translations.SetTo[LanguageNumber] + pDF_Data.SetBar + " bar").FontSize(12);
                            });
                            column.Item().Row(row =>
                            {
                                row.AutoItem().AlignCenter().Text("                                      " + pDF_Data.ElectricCutoff + Translations.ElectricCutoff[LanguageNumber] + pDF_Data.MechanicCutoff + Translations.MechanicCutoff[LanguageNumber] + pDF_Data.NoCutoff + Translations.NoCutoff[LanguageNumber]);
                            });
                            column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                            column.Item().Row(row => 
                            {
                                string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "ibnPSignatureCustomer.png");
                                row.AutoItem().AlignLeft().Column(col =>
                                {                                    
                                    col.Item().Width(310).Height(25).Border(1).AlignMiddle().Text(Translations.Customer[LanguageNumber]).FontSize(12).AlignCenter();
                                    col.Item().Width(310).Height(40).Border(1).AlignRight().Image(ImagePath_Sign_Kunde);
                                });
                                row.AutoItem().PaddingHorizontal(7).Column(col =>
                                {
                                    //TODO Datum und Ort in der Mitte einfügen vlt noch Klasse und Excel dem entsprechend erweitern
                                    DateTime SignDate = File.GetCreationTime(ImagePath_Sign_Kunde);
                                    col.Item().Text("\n" + Translations.Date[LanguageNumber] + SignDate.ToString("dd.MM.yyyy") + "\n" + Translations.Place[LanguageNumber] + pDF_Data.PlaceSignature).FontSize(12).AlignCenter();
                                });
                                row.AutoItem().AlignLeft().Column(col =>
                                {
                                    string ImagePath_Sign_Mitarbeiter = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "ibnPSignatureEmployee.png");
                                    col.Item().Width(310).Height(25).Border(1).AlignMiddle().Text(" Mitarbeiter / Employee").FontSize(12).AlignCenter();
                                    col.Item().Width(310).Height(40).Border(1).AlignRight().Image(ImagePath_Sign_Mitarbeiter);
                                });
                            });

                        });
                    });
                });
                Dokument.GeneratePdf(SavePath);
            }
            
        }
        public PDF_Data_InbetriebnahmeProtokoll GetDataForIbnP_PDF(string ExcelFilePath)
        {
            PDF_Data_InbetriebnahmeProtokoll PDF_Data_IbnP = new PDF_Data_InbetriebnahmeProtokoll();
            
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu
                    PDF_Data_IbnP.Customer = worksheet.Cells["E3"].Text;
                    PDF_Data_IbnP.ContactPerson = worksheet.Cells["E5"].Text;
                    PDF_Data_IbnP.LineConfiguration = worksheet.Cells["E7"].Text;
                    PDF_Data_IbnP.Material = worksheet.Cells["E9"].Text;
                    PDF_Data_IbnP.Viscosity = worksheet.Cells["E11"].Text;
                    PDF_Data_IbnP.FilterType = worksheet.Cells["M5"].Text;
                    PDF_Data_IbnP.SerialNumber = worksheet.Cells["M7"].Text;
                    PDF_Data_IbnP.Preloading = worksheet.Cells["M9"].Text;
                    PDF_Data_IbnP.ShimpackingLR = worksheet.Cells["M11"].Text;
                    PDF_Data_IbnP.ShimpackingZP = worksheet.Cells["O11"].Text;
                    //Prozessparameter
                    PDF_Data_IbnP.Pressure_P1.Add(worksheet.Cells["B17"].Text);
                    PDF_Data_IbnP.Pressure_P1.Add(worksheet.Cells["B19"].Text);
                    PDF_Data_IbnP.Pressure_P1.Add(worksheet.Cells["B21"].Text);
                    PDF_Data_IbnP.Pressure_P1.Add(worksheet.Cells["B23"].Text);
                    PDF_Data_IbnP.Pressure_P2.Add(worksheet.Cells["C17"].Text);
                    PDF_Data_IbnP.Pressure_P2.Add(worksheet.Cells["C19"].Text);
                    PDF_Data_IbnP.Pressure_P2.Add(worksheet.Cells["C21"].Text);
                    PDF_Data_IbnP.Pressure_P2.Add(worksheet.Cells["C23"].Text);
                    PDF_Data_IbnP.P.Add(worksheet.Cells["D17"].Text);
                    PDF_Data_IbnP.P.Add(worksheet.Cells["D19"].Text);
                    PDF_Data_IbnP.P.Add(worksheet.Cells["D21"].Text);
                    PDF_Data_IbnP.P.Add(worksheet.Cells["D23"].Text);
                    PDF_Data_IbnP.MassTemperatur.Add(worksheet.Cells["E17"].Text);
                    PDF_Data_IbnP.MassTemperatur.Add(worksheet.Cells["E19"].Text);
                    PDF_Data_IbnP.MassTemperatur.Add(worksheet.Cells["E21"].Text);
                    PDF_Data_IbnP.MassTemperatur.Add(worksheet.Cells["E23"].Text);
                    PDF_Data_IbnP.n_Extruder.Add(worksheet.Cells["F17"].Text);
                    PDF_Data_IbnP.n_Extruder.Add(worksheet.Cells["F19"].Text);
                    PDF_Data_IbnP.n_Extruder.Add(worksheet.Cells["F21"].Text);
                    PDF_Data_IbnP.n_Extruder.Add(worksheet.Cells["F23"].Text);
                    PDF_Data_IbnP.Pump.Add(worksheet.Cells["G17"].Text);
                    PDF_Data_IbnP.Pump.Add(worksheet.Cells["G19"].Text);
                    PDF_Data_IbnP.Pump.Add(worksheet.Cells["G21"].Text);
                    PDF_Data_IbnP.Pump.Add(worksheet.Cells["G23"].Text);
                    PDF_Data_IbnP.Q.Add(worksheet.Cells["H17"].Text);
                    PDF_Data_IbnP.Q.Add(worksheet.Cells["H19"].Text);
                    PDF_Data_IbnP.Q.Add(worksheet.Cells["H21"].Text);
                    PDF_Data_IbnP.Q.Add(worksheet.Cells["H23"].Text);
                    PDF_Data_IbnP.FilterElements.Add(worksheet.Cells["I17"].Text);
                    PDF_Data_IbnP.FilterElements.Add(worksheet.Cells["I19"].Text);
                    PDF_Data_IbnP.FilterElements.Add(worksheet.Cells["I21"].Text);
                    PDF_Data_IbnP.FilterElements.Add(worksheet.Cells["I23"].Text);
                    PDF_Data_IbnP.BackFlushLoss_1gr.Add(worksheet.Cells["J17"].Text);
                    PDF_Data_IbnP.BackFlushLoss_1gr.Add(worksheet.Cells["J19"].Text);
                    PDF_Data_IbnP.BackFlushLoss_1gr.Add(worksheet.Cells["J21"].Text);
                    PDF_Data_IbnP.BackFlushLoss_1gr.Add(worksheet.Cells["J23"].Text);
                    PDF_Data_IbnP.BackFlushLoss_10gr.Add(worksheet.Cells["K17"].Text);
                    PDF_Data_IbnP.BackFlushLoss_10gr.Add(worksheet.Cells["K19"].Text);
                    PDF_Data_IbnP.BackFlushLoss_10gr.Add(worksheet.Cells["K21"].Text);
                    PDF_Data_IbnP.BackFlushLoss_10gr.Add(worksheet.Cells["K23"].Text);
                    PDF_Data_IbnP.BackFlushLossInPercent.Add(worksheet.Cells["L17"].Text);
                    PDF_Data_IbnP.BackFlushLossInPercent.Add(worksheet.Cells["L19"].Text);
                    PDF_Data_IbnP.BackFlushLossInPercent.Add(worksheet.Cells["L21"].Text);
                    PDF_Data_IbnP.BackFlushLossInPercent.Add(worksheet.Cells["L23"].Text);
                    PDF_Data_IbnP.StrokeLength.Add(worksheet.Cells["M17"].Text);
                    PDF_Data_IbnP.StrokeLength.Add(worksheet.Cells["M19"].Text);
                    PDF_Data_IbnP.StrokeLength.Add(worksheet.Cells["M21"].Text);
                    PDF_Data_IbnP.StrokeLength.Add(worksheet.Cells["M23"].Text);
                    PDF_Data_IbnP.BackFlushPressure.Add(worksheet.Cells["N17"].Text);
                    PDF_Data_IbnP.BackFlushPressure.Add(worksheet.Cells["N19"].Text);
                    PDF_Data_IbnP.BackFlushPressure.Add(worksheet.Cells["N21"].Text);
                    PDF_Data_IbnP.BackFlushPressure.Add(worksheet.Cells["N23"].Text);
                    PDF_Data_IbnP.DriveForce.Add(worksheet.Cells["O17"].Text);
                    PDF_Data_IbnP.DriveForce.Add(worksheet.Cells["O19"].Text);
                    PDF_Data_IbnP.DriveForce.Add(worksheet.Cells["O21"].Text);
                    PDF_Data_IbnP.DriveForce.Add(worksheet.Cells["O23"].Text);
                    PDF_Data_IbnP.FloodingPin.Add(worksheet.Cells["P17"].Text);
                    PDF_Data_IbnP.FloodingPin.Add(worksheet.Cells["P19"].Text);
                    PDF_Data_IbnP.FloodingPin.Add(worksheet.Cells["P21"].Text);
                    PDF_Data_IbnP.FloodingPin.Add(worksheet.Cells["P23"].Text);

                    //Tabellen für die Maschienen
                    //RSF Normal
                    PDF_Data_IbnP.WStroke_Filter_RSF_1 = worksheet.Cells["D31"].Text;
                    PDF_Data_IbnP.WStroke_Filter_RSF_2 = worksheet.Cells["D32"].Text;
                    PDF_Data_IbnP.RStroke_Filter_RSF_1 = worksheet.Cells["E31"].Text;
                    PDF_Data_IbnP.RStroke_Filter_RSF_2 = worksheet.Cells["E32"].Text;
                    PDF_Data_IbnP.CycleTime_RSF_1 = worksheet.Cells["F31"].Text;
                    PDF_Data_IbnP.CycleTime_RSF_2 = worksheet.Cells["F32"].Text;
                    PDF_Data_IbnP.WStroke2_Filter_RSF_1 = worksheet.Cells["G31"].Text;
                    PDF_Data_IbnP.WStroke2_Filter_RSF_2 = worksheet.Cells["G32"].Text;
                    PDF_Data_IbnP.RStroke2_Filter_RSF_1 = worksheet.Cells["H31"].Text;
                    PDF_Data_IbnP.RStroke2_Filter_RSF_2 = worksheet.Cells["H32"].Text;
                    PDF_Data_IbnP.PPiston_Forward_1 = worksheet.Cells["I31"].Text;
                    PDF_Data_IbnP.PPiston_Forward_2 = worksheet.Cells["I32"].Text;
                    PDF_Data_IbnP.PPiston_Backward_1 = worksheet.Cells["J31"].Text;
                    PDF_Data_IbnP.PPiston_Backward_2 = worksheet.Cells["J32"].Text;
                    PDF_Data_IbnP.PPiston_Forward_2_1 = worksheet.Cells["K31"].Text;
                    PDF_Data_IbnP.PPiston_Forward_2_2 = worksheet.Cells["K32"].Text;
                    PDF_Data_IbnP.PPiston_Backward_2_1 = worksheet.Cells["L31"].Text;
                    PDF_Data_IbnP.PPiston_Backward_2_2 = worksheet.Cells["L32"].Text;
                    PDF_Data_IbnP.NumberFilterElements_1 = worksheet.Cells["M31"].Text;
                    PDF_Data_IbnP.NumberFilterElements_2 = worksheet.Cells["M32"].Text;
                    PDF_Data_IbnP.StrokesRevolt_1 = worksheet.Cells["N31"].Text;
                    PDF_Data_IbnP.StrokesRevolt_2 = worksheet.Cells["N32"].Text;
                    PDF_Data_IbnP.PuringPiston_Forward_1 = worksheet.Cells["O31"].Text;
                    PDF_Data_IbnP.PuringPiston_Forward_2 = worksheet.Cells["O32"].Text;
                    PDF_Data_IbnP.PuringPiston_Backward_1 = worksheet.Cells["P31"].Text;
                    PDF_Data_IbnP.PuringPiston_Backward_2 = worksheet.Cells["P32"].Text;

                    //SFX/SFXR
                    PDF_Data_IbnP.WStroke_Filter_SFX_1 = worksheet.Cells["D35"].Text;
                    PDF_Data_IbnP.WStroke_Filter_SFX_2 = worksheet.Cells["D36"].Text;
                    PDF_Data_IbnP.RStroke_Filter_SFX_1 = worksheet.Cells["E35"].Text;
                    PDF_Data_IbnP.RStroke_Filter_SFX_2 = worksheet.Cells["E36"].Text;
                    PDF_Data_IbnP.CycleTime_SFX_1 = worksheet.Cells["F35"].Text;
                    PDF_Data_IbnP.CycleTime_SFX_2 = worksheet.Cells["F36"].Text;
                    PDF_Data_IbnP.FloodingTime_SFX_1 = worksheet.Cells["G35"].Text;
                    PDF_Data_IbnP.FloodingTime_SFX_2 = worksheet.Cells["G36"].Text;
                    PDF_Data_IbnP.FloodingTime_Change_1 = worksheet.Cells["H35"].Text;
                    PDF_Data_IbnP.FloodingTime_Change_2 = worksheet.Cells["H36"].Text;
                    PDF_Data_IbnP.SetPressure_SFX_1 = worksheet.Cells["I35"].Text;
                    PDF_Data_IbnP.SetPressure_SFX_2 = worksheet.Cells["I36"].Text;
                    PDF_Data_IbnP.Min_Pressure_1 = worksheet.Cells["J35"].Text;
                    PDF_Data_IbnP.Min_Pressure_2 = worksheet.Cells["J36"].Text;
                    PDF_Data_IbnP.ModeOfOperation_SFX_1 = worksheet.Cells["K35"].Text;
                    PDF_Data_IbnP.ModeOfOperation_SFX_2 = worksheet.Cells["K36"].Text;
                    PDF_Data_IbnP.PreDiff_Pressure_1 = worksheet.Cells["L35"].Text;
                    PDF_Data_IbnP.PreDiff_Pressure_2 = worksheet.Cells["L36"].Text;
                    PDF_Data_IbnP.Flooding_dim_A_1 = worksheet.Cells["M35"].Text;
                    PDF_Data_IbnP.Flooding_dim_A_2 = worksheet.Cells["M36"].Text;
                    PDF_Data_IbnP.PistonCrossSection_1 = worksheet.Cells["N35"].Text;
                    PDF_Data_IbnP.PistonCrossSection_2 = worksheet.Cells["N36"].Text;
                    PDF_Data_IbnP.MeltDischarge_1 = worksheet.Cells["O35"].Text;
                    PDF_Data_IbnP.MeltDischarge_2 = worksheet.Cells["O36"].Text;

                    //KSF
                    PDF_Data_IbnP.MV_A_1 = worksheet.Cells["D39"].Text;
                    PDF_Data_IbnP.MV_A_2 = worksheet.Cells["D40"].Text;
                    PDF_Data_IbnP.MV_B_1 = worksheet.Cells["E39"].Text;
                    PDF_Data_IbnP.MV_B_2 = worksheet.Cells["E40"].Text;
                    PDF_Data_IbnP.ScreenLifeTime_1 = worksheet.Cells["F39"].Text;
                    PDF_Data_IbnP.ScreenLifeTime_2 = worksheet.Cells["F40"].Text;
                    PDF_Data_IbnP.FloodingTime_KSF_1 = worksheet.Cells["G39"].Text;
                    PDF_Data_IbnP.FloodingTime_KSF_2 = worksheet.Cells["G40"].Text;
                    PDF_Data_IbnP.Pbetween_br_Plates_1 = worksheet.Cells["H39"].Text;
                    PDF_Data_IbnP.Pbetween_br_Plates_2 = worksheet.Cells["H40"].Text;
                    PDF_Data_IbnP.Set_Pressure_KSF_1 = worksheet.Cells["I39"].Text;
                    PDF_Data_IbnP.Set_Pressure_KSF_2 = worksheet.Cells["I40"].Text;
                    PDF_Data_IbnP.Min_Pressure_KSF_1 = worksheet.Cells["J39"].Text;
                    PDF_Data_IbnP.Min_Pressure_KSF_2 = worksheet.Cells["J40"].Text;
                    PDF_Data_IbnP.Mode_Of_Operation_1 = worksheet.Cells["K39"].Text;
                    PDF_Data_IbnP.Mode_Of_Operation_2 = worksheet.Cells["K40"].Text;

                    //VIS / Korrekte Funktion der Steuerung
                    PDF_Data_IbnP.VIS = worksheet.Cells["M37"].Text;
                    PDF_Data_IbnP.dSheet = worksheet.Cells["M38"].Text;
                    PDF_Data_IbnP.KP = worksheet.Cells["M39"].Text;
                    PDF_Data_IbnP.KK = worksheet.Cells["M40"].Text;

                    if (worksheet.Cells["N38"].Text.ToUpper() == "X") {PDF_Data_IbnP.Disc_Rotation = "☑"; } else { PDF_Data_IbnP.Disc_Rotation = "☐"; }
                    if(worksheet.Cells["N39"].Text.ToUpper() == "X") { PDF_Data_IbnP.DriveLoadMeasurement = "☑"; } else { PDF_Data_IbnP.DriveLoadMeasurement = "☐"; }
                    if(worksheet.Cells["N40"].Text.ToUpper() == "X") { PDF_Data_IbnP.BackflushStrokeLength = "☑"; } else { PDF_Data_IbnP.BackflushStrokeLength = "☐"; }
                    if(worksheet.Cells["H41"].Text.ToUpper() == "X") { PDF_Data_IbnP.PhotoAttachment_Yes = "☑"; } else { PDF_Data_IbnP.PhotoAttachment_Yes = "☐"; }
                    if(worksheet.Cells["I41"].Text.ToUpper() == "X") { PDF_Data_IbnP.PhotoAttachment_No = "☑"; } else { PDF_Data_IbnP.PhotoAttachment_No = "☐"; }
                    PDF_Data_IbnP.PhotoAttachment_No_Because = worksheet.Cells["M41"].Text;

                    //Tabelle Temperaturprofil in Extrusionsrichtung
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["C45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["D45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["E45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["F45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["G45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["H45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["I45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["J45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["K45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["L45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["M45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["N45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["O45"].Text);
                    PDF_Data_IbnP.Designation_Tempprofil.Add(worksheet.Cells["P45"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["C47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["D47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["E47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["F47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["G47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["H47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["I47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["J47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["K47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["L47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["M47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["N47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["O47"].Text);
                    PDF_Data_IbnP.Temperatur_Tempprofil.Add(worksheet.Cells["P47"].Text);

                    //Questions
                    if(worksheet.Cells["A50"].Text.ToUpper() == "X") { PDF_Data_IbnP.Customer_Temperature_Meassurement_korrekt = "☑"; } else { PDF_Data_IbnP.Customer_Temperature_Meassurement_korrekt = "☐"; }
                    if(worksheet.Cells["A51"].Text.ToUpper() == "X") { PDF_Data_IbnP.PressureCutoff = "☑"; } else { PDF_Data_IbnP.PressureCutoff = "☐"; }
                    if(worksheet.Cells["A52"].Text.ToUpper() == "X") { PDF_Data_IbnP.ElectricCutoff = "☑"; } else { PDF_Data_IbnP.ElectricCutoff = "☐"; }
                    if(worksheet.Cells["F52"].Text.ToUpper() == "X") { PDF_Data_IbnP.MechanicCutoff = "☑"; } else { PDF_Data_IbnP.MechanicCutoff = "☐"; }
                    if(worksheet.Cells["J51"].Text.ToUpper() == "X") { PDF_Data_IbnP.SetTo = "☑"; } else { PDF_Data_IbnP.SetTo = "☐"; }
                    PDF_Data_IbnP.SetBar = worksheet.Cells["M51"].Text;
                    if(worksheet.Cells["J52"].Text.ToUpper() == "X") { PDF_Data_IbnP.NoCutoff = "☑"; } else { PDF_Data_IbnP.NoCutoff = "☐"; }
                }            
            return PDF_Data_IbnP;
        }

        public void Create_PDF_Of_IbnP_MRS(EngDe_For_PDF Translations, int LanguageNumber)
        {
            int NumberOfIbnP = 0;
            string ExcelFilePath = "";
            string SavePath = "";
            if (GlobalVariables.Maschiene_1 == "MRS")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_2 == "MRS")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_3 == "MRS")
            {
                NumberOfIbnP++;
            }
            if (GlobalVariables.Maschiene_4 == "MRS")
            {
                NumberOfIbnP++;
            }
            for (int i = 0; i < NumberOfIbnP; i++)
            {
                if (i == 0)
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");
                    SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.pdf");
                }
                else
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_" + (i + 1) + ".xlsx");
                    SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_" + (i + 1) + ".pdf");
                }
                
                PDF_Data_IbnP_MRS pDF_Data = GetDataForIbnP_MRS_PDF(ExcelFilePath);
                QuestPDF.Settings.License = LicenseType.Community;

                var Dokument = Document.Create(document =>
                {
                    Func<IContainer, IContainer> headerstyle = c => c
                        .Background(QuestPDF.Helpers.Colors.Grey.Lighten2)
                        .BorderColor(QuestPDF.Helpers.Colors.Grey.Darken1).Border(1)
                        .PaddingVertical(0).AlignMiddle();

                    Func<IContainer, IContainer> cellstyle = c => c
                        .Border(0.5f).BorderColor(QuestPDF.Helpers.Colors.Grey.Darken2)
                        .AlignCenter().AlignMiddle();

                    document.Page(page => 
                    {
                        page.Size(PageSizes.A4.Landscape());
                        page.Margin(10);
                        page.PageColor(QuestPDF.Helpers.Colors.White);
                        page.Header().PaddingBottom(0).BorderBottom(1).Column(column =>
                        {
                            column.Spacing(5);
                            column.Item().Row(row =>
                            {
                                row.RelativeItem().Text(Translations.ComissioningDataSheetMRS[LanguageNumber]).FontSize(20).SemiBold().AlignCenter();
                                row.ConstantItem(100)
                                .AlignRight()
                                .Image("Bilder/gneuss_png_1.png");
                            });
                            column.Item().Row(row => 
                            {
                                row.AutoItem().Column(col => 
                                {
                                    col.Item().Text(Translations.Customer[LanguageNumber] + pDF_Data.Customer).FontSize(12);
                                    col.Item().Text(Translations.ContactPerson[LanguageNumber] + pDF_Data.ContactPerson).FontSize(12);
                                    col.Item().Text(Translations.Lineconfiguration[LanguageNumber] + pDF_Data.LineConfiguration).FontSize(12);
                                    col.Item().Text("Material / Rezeptur : " + pDF_Data.Material).FontSize(12);
                                });

                                row.Spacing(250);

                                row.AutoItem().Column(col => 
                                {
                                    col.Item().Text(Translations.Extrudertype[LanguageNumber] + pDF_Data.ExtruderType).FontSize(12);
                                    col.Item().Text(Translations.OrderNo[LanguageNumber] + pDF_Data.SerialNumber).FontSize(12);
                                    col.Item().Text(Translations.FinalProduct[LanguageNumber] + pDF_Data.FinalProduct).FontSize(12);
                                    col.Item().Text(Translations.Shape[LanguageNumber] + pDF_Data.Shape).FontSize(12);
                                });
                            });
                        });
                        page.Content().Column(column => 
                        {
                            column.Item().Row(row => 
                            {
                                row.RelativeItem().Text(Translations.Processingparameters[LanguageNumber]).FontSize(16).SemiBold().AlignCenter().Underline();
                            });
                            column.Item().Table(table => 
                            {
                                table.ColumnsDefinition(columns => 
                                {
                                    columns.ConstantColumn(15);  // Nr.
                                    columns.ConstantColumn(25);  // Zeit
                                    columns.ConstantColumn(38);  // Regelung\n\nein aus
                                    columns.ConstantColumn(30);  // Pumpe
                                    columns.ConstantColumn(46);  // Auslastung\n\n%°C
                                    columns.ConstantColumn(35);  // Drehzahl \n\nsoll
                                    columns.ConstantColumn(35);  // extruder\n\n minmax
                                    columns.ConstantColumn(46);  // Auslastung\n\n%°C
                                    columns.ConstantColumn(35);  // Vakuum\n\nsoll ist
                                    columns.ConstantColumn(40);  // Viskosi\n\nViskosity 
                                    columns.ConstantColumn(40);  // meter\n\nscherung
                                    columns.ConstantColumn(36);  // MP1\n\nn.Eintrag
                                    columns.ConstantColumn(28);  // MP2\n\nn.MRS
                                    columns.ConstantColumn(38);  // MP3\n\nn.Austrag
                                    columns.ConstantColumn(36);  // MP4\n\nv.Pumpe
                                    columns.ConstantColumn(25);  // MP5\n\nDüse
                                    columns.ConstantColumn(70);  // (DoppelSpalte)Filter\nv.Filter n.Filter\n∆ P Filter
                                    columns.ConstantColumn(32);  // Sieb-\nfeinheit
                                    columns.ConstantColumn(43);  // Schnecken\n-kühlung\nist
                                    columns.ConstantColumn(35);  // Einzug\n-kühlung\nist
                                    columns.ConstantColumn(24);  // TM\nFilter
                                    columns.ConstantColumn(23);  // TM\nVisko
                                    columns.ConstantColumn(40);  // Durchsatz
                                });
                                table.Header(header => 
                                {
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.No[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Time[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Control[LanguageNumber] + "\n\n" + Translations.OnOff[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Pump[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Load[LanguageNumber] + "\n\n%/°C").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignRight().Text(Translations.ExtruderDreh[LanguageNumber] + "\n\n" + Translations.SetPoint[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignLeft().Text(Translations.Extruder_Speed[LanguageNumber] + "\n\nMin/Max").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Load[LanguageNumber] + "\n\n%/°C").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Vacuum + "\n\n" + Translations.SetAct[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignRight().Text(Translations.Viskosi[LanguageNumber] + "\n\n" + Translations.Viscosity[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignLeft().Text("meter\n\n" + Translations.ShearRate[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("MP1\n\n" + Translations.FeedZone[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("MP2\n\n" + Translations.MRS[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("MP3\n\n" + Translations.Discharge[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("MP4\n\n" + Translations.VPump[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("MP5\n\n" + Translations.Die).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("Filter\n\n" + Translations.upstrFilter[LanguageNumber] + Translations.downstrFilter[LanguageNumber] + "\n∆ P Filter").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Filter[LanguageNumber] + "\n" + Translations.Fineness[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.screw[LanguageNumber] + "\n" + Translations.Cooling[LanguageNumber] + "\n" + Translations.Actual[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.FeedZone[LanguageNumber] + "\n" + Translations.Cooling[LanguageNumber] + "\n" + Translations.Actual[LanguageNumber]).FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("TM\nFilter").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text("TM\nVisko").FontSize(8);
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Throughput[LanguageNumber]).FontSize(8);
                                });
                                for(int x = 0; x < 4; x++)
                                {
                                    table.Cell().Element(cellstyle).Text((x + 1).ToString()).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Time[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Control[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Pump[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Load[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Extruderspeed_Soll[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Extruderspeed_Min[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Load_2[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Vacuum[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Viscosimeter_Viscosity[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Viscosimeter_Shearrate[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MP1[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MP2[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MP3[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MP4[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.MP5[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Filter_P[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.FilterFineness[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Screwcooling_Actual[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Feedzone_Cooling[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.TM_Filter[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.TM_Visco[x]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.Throughput[x]).FontSize(8);
                                }                                
                            });
                            column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);
                            

                            column.Item().Text(Translations.HeatingCooling[LanguageNumber]).FontSize(16).SemiBold().AlignCenter().Underline();

                            List<string> RowHeader = new List<string>(); 
                            RowHeader.Add(Translations.Designation[LanguageNumber]);
                            RowHeader.Add(Translations.TempActualValue[LanguageNumber]);
                            RowHeader.Add(Translations.TempSetPoint[LanguageNumber]);

                            column.Item().Table(table =>
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    // Spalte 1: Breite 70 pt für den dreizeiligen Header‐Text
                                    columns.ConstantColumn(70);

                                    // Spalte 2–27: je 28 pt
                                    for (int y = 0; y < 26; y++) 
                                    { 
                                        columns.ConstantColumn(28);
                                    }
                                });
                                table.Header(header =>
                                {
                                    header.Cell().Element(headerstyle).AlignCenter().Text(Translations.Designation[LanguageNumber]).FontSize(8);
                                    for (int y = 0; y < 26; y++) 
                                    { 
                                        header.Cell().Element(headerstyle).AlignCenter().Text("HZ" + (y + 1)).FontSize(8);
                                    }
                                });
                                for (int w = 0; w < 3; w++)
                                {
                                    table.Cell().Element(cellstyle).Text(RowHeader[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ1[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ2[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ3[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ4[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ5[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ6[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ7[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ8[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ9[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ10[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ11[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ12[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ13[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ14[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ15[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ16[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ17[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ18[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ19[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ20[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ21[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ22[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ23[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ24[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ25[w]).FontSize(8);
                                    table.Cell().Element(cellstyle).Text(pDF_Data.HZ26[w]).FontSize(8);
                                }
                            });
                            column.Spacing(20);

                            column.Item().Row(row => 
                            {
                                row.RelativeItem().AlignLeft().Column(col => 
                                {
                                    col.Item().Table(table => 
                                    {
                                        table.ColumnsDefinition(columns => 
                                        {
                                            columns.ConstantColumn(100); // Spalte für die Bezeichnung
                                            columns.RelativeColumn(); // Spalte für den Wert
                                        });
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text(Translations.CoolingFeedingZone[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.Cooling_Feeding_Zone).FontSize(10);
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text(Translations.screwCooling[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.Screwcooling).FontSize(10);
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text(Translations.ChillerVacuum[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.ChillerVacuum).FontSize(10);
                                    });
                                });
                                
                                row.Spacing(20);

                                row.ConstantItem(600).AlignRight().Column(col => 
                                {
                                    col.Item().Table(table => 
                                    {
                                        table.ColumnsDefinition(columns => 
                                        {
                                            columns.ConstantColumn(100); // Spalte für die Bezeichnung
                                            columns.RelativeColumn(); // Spalte für den Wert
                                        });
                                        table.Cell().Height(96).Element(headerstyle).AlignCenter().Text(Translations.Remarks[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(96).Element(cellstyle).Text(pDF_Data.Remarks).FontSize(10);
                                    });                                    
                                });
                            });

                            column.Item().PageBreak();

                            column.Item().Text(Translations.ControlMRS[LanguageNumber]).FontSize(16).SemiBold().AlignCenter().Underline();
                            column.Item().Row(row => 
                            {
                                row.RelativeItem().Column(col => 
                                {
                                    col.Item().Table(table =>
                                    {
                                        table.ColumnsDefinition(columns =>
                                        {
                                            columns.ConstantColumn(100); // Spalte für die Bezeichnung
                                            columns.RelativeColumn(); // Spalte für den Wert
                                        });
                                        table.Header(header =>
                                        {
                                            header.Cell().ColumnSpan(2).Height(30).Element(headerstyle).AlignCenter().Text(Translations.ControlLoops[LanguageNumber]).FontSize(10);
                                        });
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text("Extruder").FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.Extruder).FontSize(10);
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text(Translations.Viscosimeter[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.Viscosimeter).FontSize(10);
                                        table.Cell().Height(32).Element(headerstyle).AlignCenter().Text(Translations.Vacuum[LanguageNumber]).FontSize(10);
                                        table.Cell().Height(32).Element(cellstyle).Text(pDF_Data.Vacuum_Control).FontSize(10);
                                    });
                                });
                                
                                row.Spacing(20);

                                row.RelativeItem().Column(col => 
                                {
                                    col.Item().Table(table => 
                                    {
                                        table.ColumnsDefinition(columns => 
                                        {
                                            columns.RelativeColumn(); // Spalte für den Wert
                                        });
                                        table.Header(header =>
                                        {
                                            header.Cell().Height(30).Element(headerstyle).AlignCenter().Text(Translations.OtherFixParameter[LanguageNumber]).FontSize(10);
                                        });
                                        table.Cell().Height(96).Element(cellstyle).Text(pDF_Data.OtherFixParameterSettings).FontSize(10);
                                    });
                                });
                            });
                            column.Item().Row(row =>
                            {
                                string ImagePath_Sign_Kunde = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "ibnPSignatureCustomer_MRS.png");
                                row.AutoItem().AlignLeft().Column(col =>
                                {
                                    col.Item().Width(310).Height(25).Border(1).AlignMiddle().Text(Translations.Customer[LanguageNumber]).FontSize(12).AlignCenter();
                                    col.Item().Width(310).Height(40).Border(1).AlignRight().Image(ImagePath_Sign_Kunde);
                                });
                                row.AutoItem().PaddingHorizontal(7).Column(col =>
                                {
                                    //TODO Datum und Ort in der Mitte einfügen vlt noch Klasse und Excel dem entsprechend erweitern                                    
                                    col.Item().Text("\n" + Translations.Date[LanguageNumber] + pDF_Data.Date_Signature + "\n" + Translations.Place[LanguageNumber] + pDF_Data.Place_Signature).FontSize(12).AlignCenter();
                                });
                                row.AutoItem().AlignLeft().Column(col =>
                                {
                                    string ImagePath_Sign_Mitarbeiter = System.IO.Path.Combine(GlobalVariables.Pfad_Unterschriften, "ibnPSignatureEmployee_MRS.png");
                                    col.Item().Width(310).Height(25).Border(1).AlignMiddle().Text(" Mitarbeiter / Employee").FontSize(12).AlignCenter();
                                    col.Item().Width(310).Height(40).Border(1).AlignRight().Image(ImagePath_Sign_Mitarbeiter);
                                });
                            });
                        });
                    });
                });
                Dokument.GeneratePdf(SavePath);
            }
        }

        public PDF_Data_IbnP_MRS GetDataForIbnP_MRS_PDF(string ExcelFilePath)
        {
            PDF_Data_IbnP_MRS PDF_Data_IbnP_MRS = new PDF_Data_IbnP_MRS();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu
                PDF_Data_IbnP_MRS.Customer = worksheet.Cells["G3"].Text;
                PDF_Data_IbnP_MRS.ContactPerson = worksheet.Cells["G4"].Text;
                PDF_Data_IbnP_MRS.LineConfiguration = worksheet.Cells["G5"].Text;
                PDF_Data_IbnP_MRS.Material = worksheet.Cells["G7"].Text;
                PDF_Data_IbnP_MRS.ExtruderType = worksheet.Cells["AF3"].Text;
                PDF_Data_IbnP_MRS.SerialNumber = worksheet.Cells["AF5"].Text;
                PDF_Data_IbnP_MRS.FinalProduct = worksheet.Cells["AF8"].Text;
                PDF_Data_IbnP_MRS.Shape = worksheet.Cells["G8"].Text;

                //Prozessparameter
                PDF_Data_IbnP_MRS.Time.Add(worksheet.Cells["C43"].Text);
                PDF_Data_IbnP_MRS.Time.Add(worksheet.Cells["C45"].Text);
                PDF_Data_IbnP_MRS.Time.Add(worksheet.Cells["C47"].Text);
                PDF_Data_IbnP_MRS.Time.Add(worksheet.Cells["C49"].Text);
                PDF_Data_IbnP_MRS.Control.Add(worksheet.Cells["E43"].Text);
                PDF_Data_IbnP_MRS.Control.Add(worksheet.Cells["E45"].Text);
                PDF_Data_IbnP_MRS.Control.Add(worksheet.Cells["E47"].Text);
                PDF_Data_IbnP_MRS.Control.Add(worksheet.Cells["E49"].Text);
                PDF_Data_IbnP_MRS.Pump.Add(worksheet.Cells["G43"].Text);
                PDF_Data_IbnP_MRS.Pump.Add(worksheet.Cells["G45"].Text);
                PDF_Data_IbnP_MRS.Pump.Add(worksheet.Cells["G47"].Text);
                PDF_Data_IbnP_MRS.Pump.Add(worksheet.Cells["G49"].Text);
                PDF_Data_IbnP_MRS.Load.Add(worksheet.Cells["I43"].Text);
                PDF_Data_IbnP_MRS.Load.Add(worksheet.Cells["I45"].Text);
                PDF_Data_IbnP_MRS.Load.Add(worksheet.Cells["I47"].Text);
                PDF_Data_IbnP_MRS.Load.Add(worksheet.Cells["I49"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Soll.Add(worksheet.Cells["K43"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Soll.Add(worksheet.Cells["K45"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Soll.Add(worksheet.Cells["K47"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Soll.Add(worksheet.Cells["K49"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Min.Add(worksheet.Cells["M43"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Min.Add(worksheet.Cells["M45"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Min.Add(worksheet.Cells["M47"].Text);
                PDF_Data_IbnP_MRS.Extruderspeed_Min.Add(worksheet.Cells["M49"].Text);
                PDF_Data_IbnP_MRS.Load_2.Add(worksheet.Cells["O43"].Text);
                PDF_Data_IbnP_MRS.Load_2.Add(worksheet.Cells["O45"].Text);
                PDF_Data_IbnP_MRS.Load_2.Add(worksheet.Cells["O47"].Text);
                PDF_Data_IbnP_MRS.Load_2.Add(worksheet.Cells["O49"].Text);
                PDF_Data_IbnP_MRS.Vacuum.Add(worksheet.Cells["Q43"].Text);
                PDF_Data_IbnP_MRS.Vacuum.Add(worksheet.Cells["Q45"].Text);
                PDF_Data_IbnP_MRS.Vacuum.Add(worksheet.Cells["Q47"].Text);
                PDF_Data_IbnP_MRS.Vacuum.Add(worksheet.Cells["Q49"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Viscosity.Add(worksheet.Cells["S43"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Viscosity.Add(worksheet.Cells["S45"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Viscosity.Add(worksheet.Cells["S47"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Viscosity.Add(worksheet.Cells["S49"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Shearrate.Add(worksheet.Cells["U43"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Shearrate.Add(worksheet.Cells["U45"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Shearrate.Add(worksheet.Cells["U47"].Text);
                PDF_Data_IbnP_MRS.Viscosimeter_Shearrate.Add(worksheet.Cells["U49"].Text);
                PDF_Data_IbnP_MRS.MP1.Add(worksheet.Cells["W43"].Text);
                PDF_Data_IbnP_MRS.MP1.Add(worksheet.Cells["W45"].Text);
                PDF_Data_IbnP_MRS.MP1.Add(worksheet.Cells["W47"].Text);
                PDF_Data_IbnP_MRS.MP1.Add(worksheet.Cells["W49"].Text);
                PDF_Data_IbnP_MRS.MP2.Add(worksheet.Cells["Y43"].Text);
                PDF_Data_IbnP_MRS.MP2.Add(worksheet.Cells["Y45"].Text);
                PDF_Data_IbnP_MRS.MP2.Add(worksheet.Cells["Y47"].Text);
                PDF_Data_IbnP_MRS.MP2.Add(worksheet.Cells["Y49"].Text);
                PDF_Data_IbnP_MRS.MP3.Add(worksheet.Cells["AA43"].Text);
                PDF_Data_IbnP_MRS.MP3.Add(worksheet.Cells["AA45"].Text);
                PDF_Data_IbnP_MRS.MP3.Add(worksheet.Cells["AA47"].Text);
                PDF_Data_IbnP_MRS.MP3.Add(worksheet.Cells["AA49"].Text);
                PDF_Data_IbnP_MRS.MP4.Add(worksheet.Cells["AC43"].Text);
                PDF_Data_IbnP_MRS.MP4.Add(worksheet.Cells["AC45"].Text);
                PDF_Data_IbnP_MRS.MP4.Add(worksheet.Cells["AC47"].Text);
                PDF_Data_IbnP_MRS.MP4.Add(worksheet.Cells["AC49"].Text);
                PDF_Data_IbnP_MRS.MP5.Add(worksheet.Cells["AE43"].Text);
                PDF_Data_IbnP_MRS.MP5.Add(worksheet.Cells["AE45"].Text);
                PDF_Data_IbnP_MRS.MP5.Add(worksheet.Cells["AE47"].Text);
                PDF_Data_IbnP_MRS.MP5.Add(worksheet.Cells["AE49"].Text);
                PDF_Data_IbnP_MRS.Filter_P.Add(worksheet.Cells["AG43"].Text);
                PDF_Data_IbnP_MRS.Filter_P.Add(worksheet.Cells["AG45"].Text);
                PDF_Data_IbnP_MRS.Filter_P.Add(worksheet.Cells["AG47"].Text);
                PDF_Data_IbnP_MRS.Filter_P.Add(worksheet.Cells["AG49"].Text);
                PDF_Data_IbnP_MRS.FilterFineness.Add(worksheet.Cells["AK43"].Text);
                PDF_Data_IbnP_MRS.FilterFineness.Add(worksheet.Cells["AK45"].Text);
                PDF_Data_IbnP_MRS.FilterFineness.Add(worksheet.Cells["AK47"].Text);
                PDF_Data_IbnP_MRS.FilterFineness.Add(worksheet.Cells["AK49"].Text);
                PDF_Data_IbnP_MRS.Screwcooling_Actual.Add(worksheet.Cells["AM43"].Text);
                PDF_Data_IbnP_MRS.Screwcooling_Actual.Add(worksheet.Cells["AM45"].Text);
                PDF_Data_IbnP_MRS.Screwcooling_Actual.Add(worksheet.Cells["AM47"].Text);
                PDF_Data_IbnP_MRS.Screwcooling_Actual.Add(worksheet.Cells["AM49"].Text);
                PDF_Data_IbnP_MRS.Feedzone_Cooling.Add(worksheet.Cells["AO43"].Text);
                PDF_Data_IbnP_MRS.Feedzone_Cooling.Add(worksheet.Cells["AO45"].Text);
                PDF_Data_IbnP_MRS.Feedzone_Cooling.Add(worksheet.Cells["AO47"].Text);
                PDF_Data_IbnP_MRS.Feedzone_Cooling.Add(worksheet.Cells["AO49"].Text);
                PDF_Data_IbnP_MRS.TM_Filter.Add(worksheet.Cells["AQ43"].Text);
                PDF_Data_IbnP_MRS.TM_Filter.Add(worksheet.Cells["AQ45"].Text);
                PDF_Data_IbnP_MRS.TM_Filter.Add(worksheet.Cells["AQ47"].Text);
                PDF_Data_IbnP_MRS.TM_Filter.Add(worksheet.Cells["AQ49"].Text);
                PDF_Data_IbnP_MRS.TM_Visco.Add(worksheet.Cells["AS43"].Text);
                PDF_Data_IbnP_MRS.TM_Visco.Add(worksheet.Cells["AS45"].Text);
                PDF_Data_IbnP_MRS.TM_Visco.Add(worksheet.Cells["AS47"].Text);
                PDF_Data_IbnP_MRS.TM_Visco.Add(worksheet.Cells["AS49"].Text);
                PDF_Data_IbnP_MRS.Throughput.Add(worksheet.Cells["AU43"].Text);
                PDF_Data_IbnP_MRS.Throughput.Add(worksheet.Cells["AU45"].Text);
                PDF_Data_IbnP_MRS.Throughput.Add(worksheet.Cells["AU47"].Text);
                PDF_Data_IbnP_MRS.Throughput.Add(worksheet.Cells["AU49"].Text);

                //Tabelle for Heating and Cooling
                PDF_Data_IbnP_MRS.HZ1.Add(worksheet.Cells["D13"].Text);
                PDF_Data_IbnP_MRS.HZ1.Add(worksheet.Cells["D15"].Text);
                PDF_Data_IbnP_MRS.HZ1.Add(worksheet.Cells["D17"].Text);
                PDF_Data_IbnP_MRS.HZ2.Add(worksheet.Cells["F13"].Text);
                PDF_Data_IbnP_MRS.HZ2.Add(worksheet.Cells["F15"].Text);
                PDF_Data_IbnP_MRS.HZ2.Add(worksheet.Cells["F17"].Text);
                PDF_Data_IbnP_MRS.HZ3.Add(worksheet.Cells["H13"].Text);
                PDF_Data_IbnP_MRS.HZ3.Add(worksheet.Cells["H15"].Text);
                PDF_Data_IbnP_MRS.HZ3.Add(worksheet.Cells["H17"].Text);
                PDF_Data_IbnP_MRS.HZ4.Add(worksheet.Cells["J13"].Text);
                PDF_Data_IbnP_MRS.HZ4.Add(worksheet.Cells["J15"].Text);
                PDF_Data_IbnP_MRS.HZ4.Add(worksheet.Cells["J17"].Text);
                PDF_Data_IbnP_MRS.HZ5.Add(worksheet.Cells["L13"].Text);
                PDF_Data_IbnP_MRS.HZ5.Add(worksheet.Cells["L15"].Text);
                PDF_Data_IbnP_MRS.HZ5.Add(worksheet.Cells["L17"].Text);
                PDF_Data_IbnP_MRS.HZ6.Add(worksheet.Cells["N13"].Text);
                PDF_Data_IbnP_MRS.HZ6.Add(worksheet.Cells["N15"].Text);
                PDF_Data_IbnP_MRS.HZ6.Add(worksheet.Cells["N17"].Text);
                PDF_Data_IbnP_MRS.HZ7.Add(worksheet.Cells["P13"].Text);
                PDF_Data_IbnP_MRS.HZ7.Add(worksheet.Cells["P15"].Text);
                PDF_Data_IbnP_MRS.HZ7.Add(worksheet.Cells["P17"].Text);
                PDF_Data_IbnP_MRS.HZ8.Add(worksheet.Cells["R13"].Text);
                PDF_Data_IbnP_MRS.HZ8.Add(worksheet.Cells["R15"].Text);
                PDF_Data_IbnP_MRS.HZ8.Add(worksheet.Cells["R17"].Text);
                PDF_Data_IbnP_MRS.HZ9.Add(worksheet.Cells["T13"].Text);
                PDF_Data_IbnP_MRS.HZ9.Add(worksheet.Cells["T15"].Text);
                PDF_Data_IbnP_MRS.HZ9.Add(worksheet.Cells["T17"].Text);
                PDF_Data_IbnP_MRS.HZ10.Add(worksheet.Cells["V13"].Text);
                PDF_Data_IbnP_MRS.HZ10.Add(worksheet.Cells["V15"].Text);
                PDF_Data_IbnP_MRS.HZ10.Add(worksheet.Cells["V17"].Text);
                PDF_Data_IbnP_MRS.HZ11.Add(worksheet.Cells["X13"].Text);
                PDF_Data_IbnP_MRS.HZ11.Add(worksheet.Cells["X15"].Text);
                PDF_Data_IbnP_MRS.HZ11.Add(worksheet.Cells["X17"].Text);
                PDF_Data_IbnP_MRS.HZ12.Add(worksheet.Cells["Z13"].Text);
                PDF_Data_IbnP_MRS.HZ12.Add(worksheet.Cells["Z15"].Text);
                PDF_Data_IbnP_MRS.HZ12.Add(worksheet.Cells["Z17"].Text);
                PDF_Data_IbnP_MRS.HZ13.Add(worksheet.Cells["AB13"].Text);
                PDF_Data_IbnP_MRS.HZ13.Add(worksheet.Cells["AB15"].Text);
                PDF_Data_IbnP_MRS.HZ13.Add(worksheet.Cells["AB17"].Text);
                PDF_Data_IbnP_MRS.HZ14.Add(worksheet.Cells["AD13"].Text);
                PDF_Data_IbnP_MRS.HZ14.Add(worksheet.Cells["AD15"].Text);
                PDF_Data_IbnP_MRS.HZ14.Add(worksheet.Cells["AD17"].Text);
                PDF_Data_IbnP_MRS.HZ15.Add(worksheet.Cells["AF13"].Text);
                PDF_Data_IbnP_MRS.HZ15.Add(worksheet.Cells["AF15"].Text);
                PDF_Data_IbnP_MRS.HZ15.Add(worksheet.Cells["AF17"].Text);
                PDF_Data_IbnP_MRS.HZ16.Add(worksheet.Cells["AH13"].Text);
                PDF_Data_IbnP_MRS.HZ16.Add(worksheet.Cells["AH15"].Text);
                PDF_Data_IbnP_MRS.HZ16.Add(worksheet.Cells["AH17"].Text);
                PDF_Data_IbnP_MRS.HZ17.Add(worksheet.Cells["AJ13"].Text);
                PDF_Data_IbnP_MRS.HZ17.Add(worksheet.Cells["AJ15"].Text);
                PDF_Data_IbnP_MRS.HZ17.Add(worksheet.Cells["AJ17"].Text);
                PDF_Data_IbnP_MRS.HZ18.Add(worksheet.Cells["AL13"].Text);
                PDF_Data_IbnP_MRS.HZ18.Add(worksheet.Cells["AL15"].Text);
                PDF_Data_IbnP_MRS.HZ18.Add(worksheet.Cells["AL17"].Text);
                PDF_Data_IbnP_MRS.HZ19.Add(worksheet.Cells["AN13"].Text);
                PDF_Data_IbnP_MRS.HZ19.Add(worksheet.Cells["AN15"].Text);
                PDF_Data_IbnP_MRS.HZ19.Add(worksheet.Cells["AN17"].Text);
                PDF_Data_IbnP_MRS.HZ20.Add(worksheet.Cells["AP13"].Text);
                PDF_Data_IbnP_MRS.HZ20.Add(worksheet.Cells["AP15"].Text);
                PDF_Data_IbnP_MRS.HZ20.Add(worksheet.Cells["AP17"].Text);
                PDF_Data_IbnP_MRS.HZ21.Add(worksheet.Cells["AR13"].Text);
                PDF_Data_IbnP_MRS.HZ21.Add(worksheet.Cells["AR15"].Text);
                PDF_Data_IbnP_MRS.HZ21.Add(worksheet.Cells["AR17"].Text);
                PDF_Data_IbnP_MRS.HZ22.Add(worksheet.Cells["AT13"].Text);
                PDF_Data_IbnP_MRS.HZ22.Add(worksheet.Cells["AT15"].Text);
                PDF_Data_IbnP_MRS.HZ22.Add(worksheet.Cells["AT17"].Text);
                PDF_Data_IbnP_MRS.HZ23.Add(worksheet.Cells["AV13"].Text);
                PDF_Data_IbnP_MRS.HZ23.Add(worksheet.Cells["AV15"].Text);
                PDF_Data_IbnP_MRS.HZ23.Add(worksheet.Cells["AV17"].Text);
                PDF_Data_IbnP_MRS.HZ24.Add(worksheet.Cells["AX13"].Text);
                PDF_Data_IbnP_MRS.HZ24.Add(worksheet.Cells["AX15"].Text);
                PDF_Data_IbnP_MRS.HZ24.Add(worksheet.Cells["AX17"].Text);
                PDF_Data_IbnP_MRS.HZ25.Add(worksheet.Cells["AZ13"].Text);
                PDF_Data_IbnP_MRS.HZ25.Add(worksheet.Cells["AZ15"].Text);
                PDF_Data_IbnP_MRS.HZ25.Add(worksheet.Cells["AZ17"].Text);
                PDF_Data_IbnP_MRS.HZ26.Add(worksheet.Cells["BB13"].Text);
                PDF_Data_IbnP_MRS.HZ26.Add(worksheet.Cells["BB15"].Text);
                PDF_Data_IbnP_MRS.HZ26.Add(worksheet.Cells["BB17"].Text);

                //Extra Info for Heating and Cooling
                PDF_Data_IbnP_MRS.Cooling_Feeding_Zone = worksheet.Cells["H20"].Text;
                PDF_Data_IbnP_MRS.Screwcooling = worksheet.Cells["H22"].Text;
                PDF_Data_IbnP_MRS.ChillerVacuum = worksheet.Cells["H24"].Text;
                PDF_Data_IbnP_MRS.Remarks = worksheet.Cells["P20"].Text;

                //Control MRS
                PDF_Data_IbnP_MRS.Extruder = worksheet.Cells["E30"].Text;
                PDF_Data_IbnP_MRS.Viscosimeter = worksheet.Cells["E33"].Text;
                PDF_Data_IbnP_MRS.Vacuum_Control = worksheet.Cells["E35"].Text;
                PDF_Data_IbnP_MRS.OtherFixParameterSettings = worksheet.Cells["AB30"].Text;

                PDF_Data_IbnP_MRS.Place_Signature = worksheet.Cells["Y53"].Text;
                PDF_Data_IbnP_MRS.Date_Signature = worksheet.Cells["Y54"].Text;
            }
            return PDF_Data_IbnP_MRS;
        }
        public string FormattedTimeSpanInHHMM(TimeSpan timeSpan)
        {//Funktion to Format TimeSpan in HH:MM format
            return Math.Truncate(timeSpan.TotalHours).ToString("00") + ":" + timeSpan.Minutes.ToString("00");
        }

        private void SendPDFToCustomer(object sender, RoutedEventArgs e)
        {
            try
            {
                // Start Outlook-Application
                Outlook.Application outlookApp = new Outlook.Application();

                // Generate new Mail Object
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Recipient
                mailItem.To = GlobalVariables.CustomerEmail;

                // Subject
                mailItem.Subject = (GlobalVariables.Sprache_Kunde == "DE") ? "Ihre Unterlagen" : "Your Documents";

                // E-Mail Body
                string bodyText = (GlobalVariables.Sprache_Kunde == "DE") ?
                    "Sehr geehrte Damen und Herren,\r\nanbei übermitteln wir Ihnen die Unterlagen zu dem durchgeführten Service-Einsatz.\r\nMit freundlichen Grüßen,\r\nGneuß Kunststofftechnik GmbH" :
                    "Dear Sirs,\r\nEnclosed you will find the documents relating to the service visit carried out.\r\nBest regards,\r\nGneuss Kunststofftechnik GmbH";

                mailItem.Body = bodyText;
                string PdfFilesPath = "";
                // Set Path Based on the Status of Online or Offline
                if (GlobalVariables.Online_or_Offline)
                {
                    PdfFilesPath = Properties.Resources.Pfad_AuftragsOrdner_On;
                }
                else
                {

                    PdfFilesPath = Properties.Resources.Pfad_AuftragsOrdner_Off;
                }

                // Replace Palceholder with actual AuftragsNR
                PdfFilesPath = string.Format(PdfFilesPath, GlobalVariables.AuftragsNR);

                // Save the PDF file name to the specified path
                string[] pdfFiles = Directory.GetFiles(PdfFilesPath, "*.pdf");

                //add all PDF files as attachments
                foreach (var pdfPath in pdfFiles)
                {
                    mailItem.Attachments.Add(pdfPath);
                }

                // Show the mail item to the user for review
                mailItem.Display();

                // Or send it directly without displaying(then comment the above line and uncomment the next line)
                // mailItem.Send();

            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler beim Senden der E-Mail: " + ex.Message);
            }
        }
    }
}

