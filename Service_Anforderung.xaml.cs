using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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
    /// Interaktionslogik für Window1.xaml
    /// </summary>
    public partial class Service_Anforderung : UserControl, IValidierbar
    {
        private bool _isInitialized = false;
        public Service_Anforderung()
        {
            InitializeComponent();

            DateTime aktuelles_Datum = DateTime.Now;
            string formatiertes_Datum = aktuelles_Datum.ToString("dd.MM.yyyy");
            dp_Druckdatum.Text = formatiertes_Datum;
            this.Loaded += ServiceAnforderungen_Loaded;
        }



        public bool HatFehlendePflichtfelder(out string fehlermeldung)
        {
            fehlermeldung = string.Empty;

            fehlermeldung = string.Empty;
            if (PrüfeMaschinenBlock(cb_Maschinentyp_1, cb_BauGröße_1, tb_MaschNr_1, out fehlermeldung))
                return true;
            if (PrüfeMaschinenBlock(cb_Maschinentyp_2, cb_BauGröße_2, tb_MaschNr_2, out fehlermeldung))
                return true;
            if (PrüfeMaschinenBlock(cb_Maschinentyp_3, cb_BauGröße_3, tb_MaschNr_3, out fehlermeldung))
                return true;
            if (PrüfeMaschinenBlock(cb_Maschinentyp_4, cb_BauGröße_4, tb_MaschNr_4, out fehlermeldung))
                return true;
            return false;
        }

        private bool PrüfeMaschinenBlock(ComboBox Maschinentyp, ComboBox Baugroeße, TextBox Maschinennummer, out string fehlermeldung)
        {
            fehlermeldung = "";

            if (Maschinentyp.SelectedItem == null)
                return false; // Block nicht aktiv → keine Pflichtprüfung

            if (Baugroeße.SelectedItem == null)
            {
                fehlermeldung = "Bitte wählen Sie eine Baugröße zu der Maschine vom Typ " + Maschinentyp.Text + "aus.";
                return true;
            }

            if (string.IsNullOrWhiteSpace(Maschinennummer.Text))
            {
                fehlermeldung = "Bitte geben Sie eine MaschinenNr zu der Maschine vom Typ " + Maschinentyp.Text + "ein";
                return true;
            }

            return false; // alles ok
        }

        private void ServiceAnforderungen_Loaded(object sender, RoutedEventArgs e)
        {
            _isInitialized = true;
        }

        private void tb_Servicetechniker_Anforderung_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            GlobalVariables.ServiceTechnicker = tb_Servicetechniker_Anforderung.Text;
        }

        private void tb_Kunde_Anforderung_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            GlobalVariables.Kunde = tb_Kunde_Anforderung.Text;
        }

        private void tb_Ansprechpartner_Anforderung_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            GlobalVariables.Ansprechpartner = tb_Ansprechpartner_Anforderung.Text;
        }

        private void tb_Anschrift_1_Anforderung_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            if (GlobalVariables.auftraginDB == false)
            {
                GlobalVariables.Anschrift_1 = tb_Anschrift_1_Anforderung.Text;
            }
        }

        private void tb_Anschrift_2_Anforderung_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;

            if (GlobalVariables.auftraginDB == false)
            {
                GlobalVariables.Anschrift_2 = tb_Anschrift_2_Anforderung.Text;
            }
        }

        private void cb_Anreise_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isInitialized) return;

            GlobalVariables.Anreise = cb_Anreise.Text;
        }
       
        private void Maschine_1_TypChanged(object sender, RoutedEventArgs e) 
        {
            //if (!_isInitialized) return;

            ComboBoxItem Item= cb_Maschinentyp_1.SelectedItem as ComboBoxItem;
            string MaschinenTyp = Item.Content.ToString();
            Baugroeßen_hinzufügen_1(MaschinenTyp);
        }
        
        private void Maschinen_vorhanden_2(object sender, RoutedEventArgs e)
        {
            //if (!_isInitialized) return;

            if (cb_Maschinentyp_2.Text != null) 
            { 
                GlobalVariables.Maschiene_2 = cb_Maschinentyp_2.Text;
                string Auftragsnummer = GlobalVariables.AuftragsNR;

                string quellOrdner =System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll.xlsm");
                string zielOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner,"Inbetriebnahme_Protokoll_2.xlsm");
                string quellOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "interner_Bericht.xlsx");
                string zielOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_2.xlsx");
                string quellOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll_MRS.xlsx");
                string zielOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_2.xlsx");

                if (!File.Exists(zielOrdner))
                {
                    File.Copy(quellOrdner, zielOrdner);
                }
                if (!File.Exists(zielOrdner2))
                {
                    File.Copy(quellOrdner2, zielOrdner2);
                }
                if (!File.Exists(zielOrdner3))
                {
                    File.Copy(quellOrdner3, zielOrdner3);
                }
                ComboBoxItem Item = cb_Maschinentyp_2.SelectedItem as ComboBoxItem;
                string MaschinenTyp = Item.Content.ToString();
                Baugroeßen_hinzufügen_2(MaschinenTyp);

            }
        }
        
        private void cb_Maschinentyp_3_TextChanged(object sender, RoutedEventArgs e)
        {
            //if (!_isInitialized) return;

            if (cb_Maschinentyp_3.Text != null)
            {
                GlobalVariables.Maschiene_3 = cb_Maschinentyp_3.Text;
                string Auftragsnummer = GlobalVariables.AuftragsNR;

                string quellOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll.xlsm");
                string zielOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_3.xlsm");
                string quellOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "interner_Bericht.xlsx");
                string zielOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_3.xlsx");
                string quellOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll_MRS.xlsx");
                string zielOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_3.xlsx");

                if (!File.Exists(zielOrdner))
                {
                    File.Copy(quellOrdner, zielOrdner);
                }
                if (!File.Exists(zielOrdner2))
                {
                    File.Copy(quellOrdner2, zielOrdner2);
                }
                if (!File.Exists(zielOrdner3))
                {
                    File.Copy(quellOrdner3, zielOrdner3);
                }

                ComboBoxItem Item = cb_Maschinentyp_3.SelectedItem as ComboBoxItem;
                string MaschinenTyp = Item.Content.ToString();
                Baugroeßen_hinzufügen_3(MaschinenTyp);
            }
        }

        private void cb_Maschinentyp_4_TextChanged(object sender, RoutedEventArgs e)
        {
            

            if (cb_Maschinentyp_4.Text != null)
            {
                GlobalVariables.Maschiene_4 = cb_Maschinentyp_4.Text;
                string Auftragsnummer = GlobalVariables.AuftragsNR;

                string quellOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll.xlsm");
                string zielOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_4.xlsm");
                string quellOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "interner_Bericht.xlsx");
                string zielOrdner2 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "interner_Bericht_4.xlsx");
                string quellOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Inbetriebnahme_Protokoll_MRS.xlsx");
                string zielOrdner3 = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Inbetriebnahme_Protokoll_MRS_4.xlsx");

                if (!File.Exists(zielOrdner))
                {
                    File.Copy(quellOrdner, zielOrdner);
                }
                if (!File.Exists(zielOrdner2))
                {
                    File.Copy(quellOrdner2, zielOrdner2);
                }
                if (!File.Exists(zielOrdner3))
                {
                    File.Copy(quellOrdner3, zielOrdner3);
                }

                ComboBoxItem Item = cb_Maschinentyp_4.SelectedItem as ComboBoxItem;
                string MaschinenTyp = Item.Content.ToString();
                Baugroeßen_hinzufügen_4(MaschinenTyp);
            }
        }

        private void Baugroeßen_hinzufügen_1(string sender)
        {
            if ( sender == "MRS")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("70");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("130");
                cb_BauGröße_1.Items.Add("160");
                cb_BauGröße_1.Items.Add("200");
            }
            else if(sender =="Jump")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("V100");
                cb_BauGröße_1.Items.Add("V600");
                cb_BauGröße_1.Items.Add("V1000");
                cb_BauGröße_1.Items.Add("V1300");
                cb_BauGröße_1.Items.Add("V2000");
                cb_BauGröße_1.Items.Add("V2800");
                cb_BauGröße_1.Items.Add("V4000");
                cb_BauGröße_1.Items.Add("V5600");
            }
            else if(sender =="RSF")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("175");
                cb_BauGröße_1.Items.Add("200");
                cb_BauGröße_1.Items.Add("250");
                cb_BauGröße_1.Items.Add("300");
                cb_BauGröße_1.Items.Add("330");
                cb_BauGröße_1.Items.Add("400");
            }
            else if(sender == "SFX")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("175");
                cb_BauGröße_1.Items.Add("200");
                cb_BauGröße_1.Items.Add("250");
                cb_BauGröße_1.Items.Add("330");
            }
            else if(sender == "SF")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("175");
                cb_BauGröße_1.Items.Add("200");
                cb_BauGröße_1.Items.Add("250");
                
            }
            else if (sender == "SFXR")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("250");
            }
            else if (sender == "KSF" || sender == "KSFx2")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("130");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("175");
                cb_BauGröße_1.Items.Add("250");
                cb_BauGröße_1.Items.Add("300");
                cb_BauGröße_1.Items.Add("350");
            }
            else if(sender == "CSF")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("175");
                cb_BauGröße_1.Items.Add("200");
                cb_BauGröße_1.Items.Add("250");
            }
            else if ( sender == "GAV")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
            }
            else if (sender == "GV")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
            }
            else if (sender == "HS" || sender =="HSS")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("20");
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("130");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("180");
                cb_BauGröße_1.Items.Add("220");
                cb_BauGröße_1.Items.Add("270");
            }
            else if (sender == "WF")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("150");
            }
            else if (sender == "WV")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("80");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("110");
                cb_BauGröße_1.Items.Add("200");
            }
            else if ( sender == "MS")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("65");
                cb_BauGröße_1.Items.Add("70");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("80");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("120");
                cb_BauGröße_1.Items.Add("150");
                cb_BauGröße_1.Items.Add("180");
                cb_BauGröße_1.Items.Add("200");
                cb_BauGröße_1.Items.Add("250");
                cb_BauGröße_1.Items.Add("254");
                cb_BauGröße_1.Items.Add("300");
                cb_BauGröße_1.Items.Add("400");
            }
            else if ( sender == "MV")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("30");
                cb_BauGröße_1.Items.Add("45");
                cb_BauGröße_1.Items.Add("60");
                cb_BauGröße_1.Items.Add("75");
                cb_BauGröße_1.Items.Add("90");
                cb_BauGröße_1.Items.Add("120");
                cb_BauGröße_1.Items.Add("150");
            }
            else if(sender == "3C-RF")
            {
                cb_BauGröße_1.Items.Clear();
                cb_BauGröße_1.Items.Add("V1000 (MRS70)");
                cb_BauGröße_1.Items.Add("V1100 (MRS090)");
                cb_BauGröße_1.Items.Add("V1300 (MRS110)");
                cb_BauGröße_1.Items.Add("V1400 (MRS130)");
                cb_BauGröße_1.Items.Add("V1500 (MRS160)");
                cb_BauGröße_1.Items.Add("V1800 (MRS200)");
            }
            else if ( sender == "VIS")
            {
                cb_BauGröße_1.Items.Clear();
            }
        }

        private void Baugroeßen_hinzufügen_2(string sender)
        {
            if (sender == "MRS")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("70");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("130");
                cb_BauGröße_2.Items.Add("160");
                cb_BauGröße_2.Items.Add("200");
            }
            else if (sender == "Jump")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("V100");
                cb_BauGröße_2.Items.Add("V600");
                cb_BauGröße_2.Items.Add("V1000");
                cb_BauGröße_2.Items.Add("V1300");
                cb_BauGröße_2.Items.Add("V2000");
                cb_BauGröße_2.Items.Add("V2800");
                cb_BauGröße_2.Items.Add("V4000");
                cb_BauGröße_2.Items.Add("V5600");
            }
            else if (sender == "RSF")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("175");
                cb_BauGröße_2.Items.Add("200");
                cb_BauGröße_2.Items.Add("250");
                cb_BauGröße_2.Items.Add("300");
                cb_BauGröße_2.Items.Add("330");
                cb_BauGröße_2.Items.Add("400");
            }
            else if (sender == "SFX")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("175");
                cb_BauGröße_2.Items.Add("200");
                cb_BauGröße_2.Items.Add("250");
                cb_BauGröße_2.Items.Add("330");
            }
            else if (sender == "SF")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("175");
                cb_BauGröße_2.Items.Add("200");
                cb_BauGröße_2.Items.Add("250");

            }
            else if (sender == "SFXR")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("250");
            }
            else if (sender == "KSF" || sender == "KSFx2")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("130");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("175");
                cb_BauGröße_2.Items.Add("250");
                cb_BauGröße_2.Items.Add("300");
                cb_BauGröße_2.Items.Add("350");
            }
            else if (sender == "CSF")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("175");
                cb_BauGröße_2.Items.Add("200");
                cb_BauGröße_2.Items.Add("250");
            }
            else if (sender == "GAV")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
            }
            else if (sender == "GV")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
            }
            else if (sender == "HS")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("20");
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("130");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("220");
                cb_BauGröße_2.Items.Add("270");
            }
            else if (sender == "WF")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("150");
            }
            else if (sender == "WV")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("80");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("110");
                cb_BauGröße_2.Items.Add("200");
            }
            else if (sender == "MS")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("65");
                cb_BauGröße_2.Items.Add("70");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("80");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("120");
                cb_BauGröße_2.Items.Add("150");
                cb_BauGröße_2.Items.Add("180");
                cb_BauGröße_2.Items.Add("200");
                cb_BauGröße_2.Items.Add("250");
                cb_BauGröße_2.Items.Add("254");
                cb_BauGröße_2.Items.Add("300");
                cb_BauGröße_2.Items.Add("400");
            }
            else if (sender == "MV")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("30");
                cb_BauGröße_2.Items.Add("45");
                cb_BauGröße_2.Items.Add("60");
                cb_BauGröße_2.Items.Add("75");
                cb_BauGröße_2.Items.Add("90");
                cb_BauGröße_2.Items.Add("120");
                cb_BauGröße_2.Items.Add("150");
            }
            else if (sender == "3C-RF")
            {
                cb_BauGröße_2.Items.Clear();
                cb_BauGröße_2.Items.Add("V1000 (MRS70)");
                cb_BauGröße_2.Items.Add("V1100 (MRS090)");
                cb_BauGröße_2.Items.Add("V1300 (MRS110)");
                cb_BauGröße_2.Items.Add("V1400 (MRS130)");
                cb_BauGröße_2.Items.Add("V1500 (MRS160)");
                cb_BauGröße_2.Items.Add("V1800 (MRS200)");
            }
            else if (sender == "VIS")
            {
                cb_BauGröße_2.Items.Clear();
            }
        }

        private void Baugroeßen_hinzufügen_3(string sender)
        {
            if (sender == "MRS")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("70");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("130");
                cb_BauGröße_3.Items.Add("160");
                cb_BauGröße_3.Items.Add("200");
            }
            else if (sender == "Jump")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("V100");
                cb_BauGröße_3.Items.Add("V600");
                cb_BauGröße_3.Items.Add("V1000");
                cb_BauGröße_3.Items.Add("V1300");
                cb_BauGröße_3.Items.Add("V2000");
                cb_BauGröße_3.Items.Add("V2800");
                cb_BauGröße_3.Items.Add("V4000");
                cb_BauGröße_3.Items.Add("V5600");
            }
            else if (sender == "RSF")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("175");
                cb_BauGröße_3.Items.Add("200");
                cb_BauGröße_3.Items.Add("250");
                cb_BauGröße_3.Items.Add("300");
                cb_BauGröße_3.Items.Add("330");
                cb_BauGröße_3.Items.Add("400");
            }
            else if (sender == "SFX")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("175");
                cb_BauGröße_3.Items.Add("200");
                cb_BauGröße_3.Items.Add("250");
                cb_BauGröße_3.Items.Add("330");
            }
            else if (sender == "SF")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("175");
                cb_BauGröße_3.Items.Add("200");
                cb_BauGröße_3.Items.Add("250");

            }
            else if (sender == "SFXR")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("250");
            }
            else if (sender == "KSF" || sender == "KSFx2")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("130");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("175");
                cb_BauGröße_3.Items.Add("250");
                cb_BauGröße_3.Items.Add("300");
                cb_BauGröße_3.Items.Add("350");
            }
            else if (sender == "CSF")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("175");
                cb_BauGröße_3.Items.Add("200");
                cb_BauGröße_3.Items.Add("250");
            }
            else if (sender == "GAV")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
            }
            else if (sender == "GV")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
            }
            else if (sender == "HS")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("20");
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("130");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("220");
                cb_BauGröße_3.Items.Add("270");
            }
            else if (sender == "WF")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("150");
            }
            else if (sender == "WV")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("80");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("110");
                cb_BauGröße_3.Items.Add("200");
            }
            else if (sender == "MS")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("65");
                cb_BauGröße_3.Items.Add("70");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("80");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("120");
                cb_BauGröße_3.Items.Add("150");
                cb_BauGröße_3.Items.Add("180");
                cb_BauGröße_3.Items.Add("200");
                cb_BauGröße_3.Items.Add("250");
                cb_BauGröße_3.Items.Add("254");
                cb_BauGröße_3.Items.Add("300");
                cb_BauGröße_3.Items.Add("400");
            }
            else if (sender == "MV")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("30");
                cb_BauGröße_3.Items.Add("45");
                cb_BauGröße_3.Items.Add("60");
                cb_BauGröße_3.Items.Add("75");
                cb_BauGröße_3.Items.Add("90");
                cb_BauGröße_3.Items.Add("120");
                cb_BauGröße_3.Items.Add("150");
            }
            else if (sender == "3C-RF")
            {
                cb_BauGröße_3.Items.Clear();
                cb_BauGröße_3.Items.Add("V1000 (MRS70)");
                cb_BauGröße_3.Items.Add("V1100 (MRS090)");
                cb_BauGröße_3.Items.Add("V1300 (MRS110)");
                cb_BauGröße_3.Items.Add("V1400 (MRS130)");
                cb_BauGröße_3.Items.Add("V1500 (MRS160)");
                cb_BauGröße_3.Items.Add("V1800 (MRS200)");
            }
            else if (sender == "VIS")
            {
                cb_BauGröße_3.Items.Clear();
            }
        }

        private void Baugroeßen_hinzufügen_4(string sender)
        {
            if (sender == "MRS")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("70");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("130");
                cb_BauGröße_4.Items.Add("160");
                cb_BauGröße_4.Items.Add("200");
            }
            else if (sender == "Jump")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("V100");
                cb_BauGröße_4.Items.Add("V600");
                cb_BauGröße_4.Items.Add("V1000");
                cb_BauGröße_4.Items.Add("V1300");
                cb_BauGröße_4.Items.Add("V2000");
                cb_BauGröße_4.Items.Add("V2800");
                cb_BauGröße_4.Items.Add("V4000");
                cb_BauGröße_4.Items.Add("V5600");
            }
            else if (sender == "RSF")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("175");
                cb_BauGröße_4.Items.Add("200");
                cb_BauGröße_4.Items.Add("250");
                cb_BauGröße_4.Items.Add("300");
                cb_BauGröße_4.Items.Add("330");
                cb_BauGröße_4.Items.Add("400");
            }
            else if (sender == "SFX")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("175");
                cb_BauGröße_4.Items.Add("200");
                cb_BauGröße_4.Items.Add("250");
                cb_BauGröße_4.Items.Add("330");
            }
            else if (sender == "SF")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("175");
                cb_BauGröße_4.Items.Add("200");
                cb_BauGröße_4.Items.Add("250");

            }
            else if (sender == "SFXR")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("250");
            }
            else if (sender == "KSF" || sender == "KSFx2")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("130");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("175");
                cb_BauGröße_4.Items.Add("250");
                cb_BauGröße_4.Items.Add("300");
                cb_BauGröße_4.Items.Add("350");
            }
            else if (sender == "CSF")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("175");
                cb_BauGröße_4.Items.Add("200");
                cb_BauGröße_4.Items.Add("250");
            }
            else if (sender == "GAV")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
            }
            else if (sender == "GV")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
            }
            else if (sender == "HS")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("20");
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("130");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("220");
                cb_BauGröße_4.Items.Add("270");
            }
            else if (sender == "WF")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("150");
            }
            else if (sender == "WV")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("80");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("110");
                cb_BauGröße_4.Items.Add("200");
            }
            else if (sender == "MS")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("65");
                cb_BauGröße_4.Items.Add("70");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("80");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("120");
                cb_BauGröße_4.Items.Add("150");
                cb_BauGröße_4.Items.Add("180");
                cb_BauGröße_4.Items.Add("200");
                cb_BauGröße_4.Items.Add("250");
                cb_BauGröße_4.Items.Add("254");
                cb_BauGröße_4.Items.Add("300");
                cb_BauGröße_4.Items.Add("400");
            }
            else if (sender == "MV")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("30");
                cb_BauGröße_4.Items.Add("45");
                cb_BauGröße_4.Items.Add("60");
                cb_BauGröße_4.Items.Add("75");
                cb_BauGröße_4.Items.Add("90");
                cb_BauGröße_4.Items.Add("120");
                cb_BauGröße_4.Items.Add("150");
            }
            else if (sender == "3C-RF")
            {
                cb_BauGröße_4.Items.Clear();
                cb_BauGröße_4.Items.Add("V1000 (MRS70)");
                cb_BauGröße_4.Items.Add("V1100 (MRS090)");
                cb_BauGröße_4.Items.Add("V1300 (MRS110)");
                cb_BauGröße_4.Items.Add("V1400 (MRS130)");
                cb_BauGröße_4.Items.Add("V1500 (MRS160)");
                cb_BauGröße_4.Items.Add("V1800 (MRS200)");
            }
            else if (sender == "VIS")
            {
                cb_BauGröße_4.Items.Clear();
            }
        }

        //Funktion zum Überprüfen der Maschinen-Nr. auf 4 Ziffern wenn ein Maschinentyp ausgewählt wurde
        private void tb_MaNr1_LostFocus(object sender, RoutedEventArgs e)
        {
            if (cb_Maschinentyp_1.SelectedItem == null) return;
            CheckIfLegitMaschineNo(sender);
        }
        private void tb_MaNr2_LostFocus(object sender, RoutedEventArgs e)
        {
            if (cb_Maschinentyp_2.SelectedItem == null) return;
            CheckIfLegitMaschineNo(sender);
        }
        private void tb_MaNr3_LostFocus(object sender, RoutedEventArgs e)
        {
            if (cb_Maschinentyp_3.SelectedItem == null) return;
            CheckIfLegitMaschineNo(sender);
        }
        private void tb_MaNr4_LostFocus(object sender, RoutedEventArgs e)
        {
            if (cb_Maschinentyp_4.SelectedItem == null) return;
            CheckIfLegitMaschineNo(sender);
        }

        private void CheckIfLegitMaschineNo(object sender)
        {
            //Check if the UserControl is initialized
            if (!_isInitialized) return;

            if (sender is TextBox tb)
            {
                //Check if TB is only digits and has a length of 4
                if (tb.Text.All(char.IsDigit) && tb.Text.Length != 4)
                {
                    MessageBox.Show("Die Maschinen-Nr. ist ungültig. sollte nur aus Zahlen bestehen und 4 Zeichen lang sein. bei einer Länge von 3 Zahlen trage eine 0 vorne ein.");
                    tb.Text = "";
                }
            }
        }

    }
    
}
