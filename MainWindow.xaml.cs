using OfficeOpenXml;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Threading;
using System.Windows.Media.TextFormatting;
using System.Runtime.DesignerServices;
using System.Data.SqlClient;
using System.Windows.Controls.Primitives;
using System.Net.NetworkInformation;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using static ServiceTool.MainWindow;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime.Workdays;

//TODO-List
// Deutsch Englisch gegenenfalls für den rest auch noch einbauen
//Anreise spalte Stundennachweis anpassen
//Checkradiobuttons funktion rausnehmen



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

            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;

            this.Width = screenWidth * 0.95;
            this.Height = screenHeight * 0.95;
            
            this.Loaded += MainWindow_Loaded;

            SaveCellMapping_InDictionarys();

            List<string> Lbl_Names = new List<string>();
            List<string> Lbl_Content_German = new List<string>();
            List<string> Lbl_Content_English = new List<string>();

            GetLabelContent();


            bool File_Connection_Test = IstServerErreichbar(Properties.Resources.IP_File02);
            bool DB_Connection_Test = IstServerErreichbar(Properties.Resources.IP_SQL04);

            if (File_Connection_Test && DB_Connection_Test)
            {
                collect_Data_From_Database();
                GlobalVariables.Online_or_Offline = true;
                lbl_OnlineOfflineAnzeige.Content = "Online";
                lbl_OnlineOfflineAnzeige.Background = Brushes.Green;
            }
            else
            {
                GlobalVariables.Online_or_Offline = false;
                lbl_OnlineOfflineAnzeige.Content = "Offline";
                lbl_OnlineOfflineAnzeige.Background = Brushes.Red;
            }

            CC.Content = new Startseite();
            //CC.Content = new Inbetriebnahme_Protokoll();
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

            string json = File.ReadAllText(Properties.Resources.Path_LanguageJson_IBNP);
            List<SprachtabelleEntry> sprachtabelleEntries = JsonConvert.DeserializeObject<List<SprachtabelleEntry>>(json);

            foreach(SprachtabelleEntry entry in sprachtabelleEntries)
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

            if(Dokument is FrameworkElement element) { 
                foreach (string lblName in sprachtabelle.Lbl_Names)
                {
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
            string Connectionstring = Properties.Resources.Connectionstring;           

            string DB_Query = Properties.Resources.DB_Abfrage;

            using (SqlConnection connection = new SqlConnection(Connectionstring)) 
            {
                SqlDataAdapter adapter = new SqlDataAdapter(DB_Query, connection);
                GlobalVariables.dt = new DataTable();
                adapter.Fill(GlobalVariables.dt);
            }
            // Ausgabe der Spaltennamen
            foreach (DataColumn column in GlobalVariables.dt.Columns)
            {
                Console.Write($"{column.ColumnName}\t");
            }
            Console.WriteLine();

            // Ausgabe der Zeilen
            foreach (DataRow row in GlobalVariables.dt.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    Console.Write($"{item}\t");
                }
                Console.WriteLine();
            }
        }

        private void rbt_Startseite_Checked(object sender, RoutedEventArgs e) // wenn der  radiobutton von der Startseite angehackt wird, wird die Startseite erstellt und im Content Control platziert
        {
            if (_blockiereUControlWechsel) return;
            CC.Content = new Startseite();

        }

        private void rbt_ServiceAnforderung_Checked(object sender, RoutedEventArgs e)
        {
            
            var sa = new Service_Anforderung();
            CC.Content = sa;
                        
            string Auftragsnummer = GlobalVariables.AuftragsNR;                      

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner,"Service_Anforderungen.xlsx");

            Laden(ExcelFilePath, "Serviceanforderungen");

            sa.tb_Auftragsnummer.Text = GlobalVariables.AuftragsNR;

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
        private void rbt_ServiceAnforderung_UnChecked(object sender, RoutedEventArgs e) 
        {
            var sa = CC.Content as Service_Anforderung;

            if (_isInitialized)
            {

                if (sa is IValidierbar validierbar)
                {
                    if (validierbar.HatFehlendePflichtfelder(out string fehlermeldung))
                    {
                        MessageBox.Show(fehlermeldung, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);

                        _blockiereUControlWechsel = true;
                        
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

            speichern(ExcelFilePath, "Serviceanforderungen");

            GlobalVariables.Kunde = sa.tb_End_Kunde.Text;
            GlobalVariables.Ansprechpartner = sa.tb_Ansprechpartner_Anforderung.Text;
            GlobalVariables.Anschrift_1 = sa.tb_Anschrift_1_Anforderung.Text;
            GlobalVariables.Anschrift_2 = sa.tb_Anschrift_2_Anforderung.Text;
            GlobalVariables.Anreise = sa.cb_Anreise.Text;
            GlobalVariables.ServiceTechnicker = sa.tb_Servicetechniker_Anforderung.Text;

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
            if (sa.dp_Besuchsdatum_Start.SelectedDate != null && sa.dp_Besuchsdatum_Ende.SelectedDate != null)
            {
                GlobalVariables.StartServiceEinsatz = (DateTime)sa.dp_Besuchsdatum_Start.SelectedDate;
                GlobalVariables.EndeServiceEinsatz = (DateTime)sa.dp_Besuchsdatum_Ende.SelectedDate;
            }
        }

        private void rbt_Stundennachweis_Checked(object sender, RoutedEventArgs e)
        {
            if (_blockiereUControlWechsel) return;
            // Hier wird der Stundennachweis geladen
            var sn = new Stundennachweis();
            CC.Content = sn;

            //Textboxen die aus den Service anforderungen übernommen werden
            

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
                    if(GlobalVariables.Anreise != "")
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

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            

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

            ibnP.tb_Kunde_ibnProtokoll.Text = GlobalVariables.Kunde;
            ibnP.tb_Ansprechpartner_ibnProtokoll.Text = GlobalVariables.Ansprechpartner;
            ibnP.tb_KundeMaterial_ibnProtokoll.Text = GlobalVariables.Material;
            
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


            isFirstLoad = false;
        }//Ende InbP Laden        
                                 
        public void InbetriebnahmeProtokoll_Speichern(string lastSelectedSite)
        {
            if (_blockiereUControlWechsel) return;
            var ibnP = CC.Content as Inbetriebnahme_Protokoll;

            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string ImagePath_Sign_Kunde = $@"C:\Users\jgadmin\Documents\Service aufträge\{Auftragsnummer}\Anhänge\Unterschriften\ibnPSignatureCustomer.png";

            string ImagePath_Sign_Technican = $@"C:\Users\jgadmin\Documents\Service aufträge\{Auftragsnummer}\Anhänge\Unterschriften\ibnPSignatureEmployee.png";

            string ExcelFilePath = "";

            // da es die möglichkeit mehrer IbnP gibt muss überprüft werden welche aktuell bearbeitet wurde an dem Punkt wo der  IbnP Radiobutton abgehackt wurde
            if (lastSelectedSite == "" || lastSelectedSite == GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll.xlsm");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_1 + " " + GlobalVariables.Baugroeße_1;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_1;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_2.xlsm");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_2 + " " + GlobalVariables.Baugroeße_2;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_2;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_3 + " " + GlobalVariables.Baugroeße_3;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_3;
            }
            else if (lastSelectedSite == GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4)
            {
                ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                ibnP.tb_Filtertyp_ibnProtokoll.Text = GlobalVariables.Maschiene_4 + " " + GlobalVariables.Baugroeße_4;
                ibnP.tb_SerienNr_ibnProtokoll.Text = GlobalVariables.MaschinenNr_4;
            }

            speichern(ExcelFilePath, "IbnP");

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

            string ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS.xlsx");

            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext= LicenseContext.NonCommercial;

                using(var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                {
                    var Worksheet = package.Workbook.Worksheets[0];

                    speichern(ExcelFilePath, "IbnP_MRS");

                }
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
            string Auftragsnummer = GlobalVariables.AuftragsNR;

            string Pfad_fuerAnhaenge = GlobalVariables.Pfad_Anhaenge;

            Process.Start("explorer.exe", Pfad_fuerAnhaenge);
        } // Funktion um mit einem Button klick in den Anhang ordner des Auftrags zu gelangen

        public void SaveSignatureAsImage(InkCanvas inkCanvas, string filePath)
        {
            // 1. Layout aktualisieren, damit Größen und Striche definitiv bereitstehen
            inkCanvas.UpdateLayout();  // sicherstellen, dass ActualWidth/Height korrekt:contentReference[oaicite:10]{index=10}

            int width = (int)inkCanvas.ActualWidth;
            int height = (int)inkCanvas.ActualHeight;
            if (width == 0 || height == 0) return; // InkCanvas nicht sichtbar oder keine Größe

            // 2. DrawingVisual erzeugen und darin das InkCanvas "nachmalen"
            var dv = new DrawingVisual();
            using (DrawingContext dc = dv.RenderOpen())
            {
                // (Optional) Hintergrund zeichnen, falls InkCanvas einen Hintergrund hat:
                if (inkCanvas.Background != null)
                {
                    // Hintergrund als Brush füllen (z.B. Farbe) über die ganze Fläche
                    dc.DrawRectangle(inkCanvas.Background, null, new Rect(0, 0, width, height));
                }
                // Alle Striche zeichnen – entweder einzeln oder gesamte StrokeCollection:
                // Variante A: Alle Striche einzeln zeichnen
                foreach (System.Windows.Ink.Stroke stroke in inkCanvas.Strokes)
                {
                    stroke.Draw(dc);  // Stroke zeichnet sich selbst mit seinen DrawingAttributes
                }
                // Variante B (alternative): inkCanvas.Strokes.Draw(dc);
            } // DrawingContext auto-close here

            // 3. RenderTargetBitmap mit passendem PixelFormat anlegen und DrawingVisual rendern
            var rtb = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);

            // 4. Als PNG-Datei speichern
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(rtb));
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                encoder.Save(fs);
            }
        }

        public void speichern(string ExcelFilePath, string Seite)
        {          

            if (File.Exists(ExcelFilePath))
            {
                
                using (var package = new ExcelPackage(ExcelFilePath))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    object Dokument = null; // Dokument in dem das Aktuell offene Formular gespeichert wird

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

                    foreach (KeyValuePair<string,string> CellMapping in CellMappings) // Schleife über die Länge des ZellenObjekte Arrays
                    {
                        string Zelle = CellMapping.Key;

                        string Objectname = CellMapping.Value;

                        string Object_bezeichnung = Objectname.Substring(0, 2).ToUpper(); // hier werden die ersten zwei Buchstaben der Object namen abgetrennt da man anhand dessen die Objekttypen unterscheiden kann

                        
                        if (Dokument is FrameworkElement element)
                        {

                            switch (Object_bezeichnung)
                            {
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
        } // allgemeine Speicherfunktion die auf jedes Dokument anwendbar ist

        public void Laden(string ExcelFilePath, string Seite)
        {           

            if (File.Exists(ExcelFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
            string Pfad_DokumentOrdner = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string Pfad_Servicetool_Lokal = string.Format(Properties.Resources.Pfad_AuftragsOrdner_Off, GlobalVariables.AuftragsNR);
            string Lokaler_Pfad = System.IO.Path.Combine(Pfad_DokumentOrdner, Pfad_Servicetool_Lokal);
            
            if (Directory.Exists(Lokaler_Pfad))
            {
                foreach (string datei in Directory.GetFiles(GlobalVariables.Pfad_AuftragsOrdner))
                {
                    string DateiName = System.IO.Path.GetFileName(datei);
                    string ZielDateiPfad = System.IO.Path.Combine(Lokaler_Pfad, DateiName);

                    if (!File.Exists(ZielDateiPfad))
                    {
                        File.Copy(datei, ZielDateiPfad);
                    }
                }
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

        }

        private void Creat_PDF_Dokuments(object sender, EventArgs e) 
        {
            MessageBox.Show("Stelle sicher das alle Dokumente vollständig ausgefüllt sind!");
            
            //Hier werden alle PDF ersellungen getrigerd
            Create_PDF_Of_Stundennachweis();
        }

        private void Create_PDF_Of_Stundennachweis()
        {
            Stundennachweis_PDF_Data pDF_Data = GetDataForPDF();
            QuestPDF.Settings.License = LicenseType.Community;
            string SavePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.pdf");

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(35);
                    page.Size(PageSizes.A4);
                    page.PageColor(QuestPDF.Helpers.Colors.White);
                    page.Header()
                    .PaddingBottom(10)
                    .BorderBottom(1)
                    .Column(column =>
                    {
                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Stundennachweis").FontSize(27).Bold();

                            row.ConstantItem(100)
                            .AlignRight()
                            .Image("Bilder/gneuss_png_1.png");

                        });

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Column(col=>
                            {
                                col.Item().Text("Auftragsnummer : " + GlobalVariables.AuftragsNR).FontSize(12);
                                col.Item().Text("Kunde/Customer : " + pDF_Data.Customer).FontSize(12);
                                col.Item().Text("Contact person : " + pDF_Data.ContactPerson).FontSize(12);
                            });


                            row.RelativeItem().AlignRight().Column(col =>
                            {
                                col.Item().Text("Datum : " + DateTime.Now.ToString("dd.MM.yyyy")).FontSize(12);
                                col.Item().Text("Techniker : " + pDF_Data.ServiceTechnician).FontSize(12);
                                col.Item().Text("Anschrift : " + pDF_Data.Adress1).FontSize(12);
                                col.Item().Text("                 " + pDF_Data.Adress2).FontSize(12);
                            });                            
                        });
                        
                    });
                    page.Content()
                    .Column(column =>
                    {
                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Reisezeit/ Traveltime").FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Verkehrsmittel/Means of Transport : " + pDF_Data.meansofTransport).FontSize(12).AlignCenter();
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
                                header.Cell().Element(headerstyle).Text("Datum/Date Start").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Time Start").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Datum/Date End").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Time End").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Pause/Break").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Total Time").FontSize(14);
                                header.Cell().Element(headerstyle).Text("Kilometer").FontSize(14);
                            });

                            

                            table.Cell().Element(cellstyle).Text("Anreise").FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateArrivalStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeArrivalStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateArrivalEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeArrivalEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.BreakArrival).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalTimeArrival).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalKilometersArrival).FontSize(12);

                            table.Cell().Element(cellstyle).Text("Abfahrt").FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateDepartureStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeDepartureStart).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.DateDepartureEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TimeDepartureEnd).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.BreakDeparture).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalTimeDeparture).FontSize(12);
                            table.Cell().Element(cellstyle).Text(pDF_Data.TotalKilometersDeparture).FontSize(12);
                                                       
                        });
                        column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Spacing(5);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Arbeitszeit / Working time ").FontSize(20).AlignCenter();
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
                                header.Cell().Element(headerstyle).Text("Datum/Date").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Time Start").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Time End").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Pause/Break").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Note").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Normal std.").FontSize(12);
                                header.Cell().Element(headerstyle).Text("overtime").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Nightwork").FontSize(12);
                                header.Cell().Element(headerstyle).Text("Total Time").FontSize(12);
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

                        column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Spacing(15);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text(" Bericht / report ").FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);
                        column.Item().Row(row =>
                        {
                            int countReports = pDF_Data.Report.Count;
                            string allReports = "";
                            for (int i = 0; i < countReports; i++)
                            {
                                allReports += $"Woche " + (i+1) + ":\n" + pDF_Data.Report[i] + "\n";
                            }
                            row.RelativeItem().Text(allReports);
                        });

                        column.Item().PageBreak();

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Schulung / Training").FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);


                        column.Item().Row(row => {
                            row.AutoItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.SetupAufbau + " Setup/Aufbau").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.OperatingPrinciple + " Operating principle/Funktionsweise").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.BriefingControlSystem + " Abnahme/Acceptance").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.Sonstiges + " Schulung/Training").FontSize(10).AlignLeft();
                            });

                            row.Spacing(50);

                            row.RelativeItem().Column(col =>
                            {
                                col.Item().Text(pDF_Data.Troubleshooting + " Troubleshooting/Störungsbeseitigung").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.OperationOfWholeEquipment + " Operating of whole Equipment/Bedienung aller Anlagenteile").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.SafetyInstructions + " Safety instructions/Acceptance").FontSize(10).AlignLeft();
                                col.Item().Text(pDF_Data.Sonstiges + " Maintenance/Wartung").FontSize(10).AlignLeft();
                            });
                        });

                        column.Item().PaddingVertical(5).LineHorizontal(1).LineColor(QuestPDF.Helpers.Colors.Black);

                        column.Item().Row(row =>
                        {
                            row.RelativeItem().Text("Evaluation / Bewertung").FontSize(20).AlignCenter();
                        });

                        column.Spacing(5);
                        //"😊", "😐", "🙁"

                        column.Item().Row(row =>
                        {
                            row.AutoItem().Column(col =>
                            { 
                                col.Item().Text("Wie zufrieden sind Sie mit dem Service?").FontSize(10).AlignLeft();
                                col.Item().Text("How satisfied are you with the service?").FontSize(10).AlignLeft();
                                col.Item().Text(" ").FontSize(10).AlignLeft();
                                col.Item().Text("Wie zufrieden sind Sie mit dem Gneuß-Support?").FontSize(10).AlignLeft();
                                col.Item().Text("How satisfied are you with the Gneuß support?").FontSize(10).AlignLeft();
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
                    });
                });
            });
            document.GeneratePdf(SavePath);

        }

        public string FormattedTimeSpanInHHMM(TimeSpan timeSpan)
        {
            return Math.Truncate(timeSpan.TotalHours).ToString("00") + ":" + timeSpan.Minutes.ToString("00");
        }

        public Stundennachweis_PDF_Data GetDataForPDF()
        {
            TimeSpan ServiceDurationInDays = GlobalVariables.EndeServiceEinsatz - GlobalVariables.StartServiceEinsatz;

            double weeksnotRounded = ServiceDurationInDays.TotalDays / 7;
            int NumberOfStundennachweis = (int)Math.Ceiling(weeksnotRounded);

            TimeSpan TotalNormalStd = new TimeSpan();
            TimeSpan TotalOverTime = new TimeSpan();
            TimeSpan TotalNightwork = new TimeSpan();
            TimeSpan TotalTime = new TimeSpan();


            Stundennachweis_PDF_Data PDF_Data = new Stundennachweis_PDF_Data();
            for (int i = 0; i < NumberOfStundennachweis; i++)
            {
                string ExcelFilePath = "";
                string PDFFilePath = "";

                

                if (i == 0)
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm");
                    PDFFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.pdf");
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu

                        //Werte für den Header der PDF abspeichern
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

                        //☑☐
                        TotalNormalStd += TimeSpan.Parse(worksheet.Cells["J31"].Text);
                        TotalOverTime += TimeSpan.Parse(worksheet.Cells["M31"].Text);
                        TotalNightwork += TimeSpan.Parse(worksheet.Cells["O31"].Text);
                        TotalTime += TimeSpan.Parse(worksheet.Cells["Q31"].Text);

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
                {
                    ExcelFilePath = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_" + (i + 1) + ".xlsm");
                    using (var package = new ExcelPackage(new FileInfo(ExcelFilePath)))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // Greife auf das erste Arbeitsblatt zu
                        PDF_Data.Report.Add(worksheet.Cells["A35"].Text);

                        TotalNormalStd += TimeSpan.Parse(worksheet.Cells["J31"].Text);
                        TotalOverTime += TimeSpan.Parse(worksheet.Cells["M31"].Text);
                        TotalNightwork += TimeSpan.Parse(worksheet.Cells["O31"].Text);
                        TotalTime += TimeSpan.Parse(worksheet.Cells["Q31"].Text);
                        
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
            return PDF_Data;
        }

    }
}

