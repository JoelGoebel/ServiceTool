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

namespace ServiceTool
{
    /// <summary>
    /// Interaktionslogik für Window2.xaml
    /// </summary>
    public partial class Stundennachweis : UserControl
    {
        private bool _isInitialized = false;
        public Stundennachweis()
        {
            InitializeComponent();
            TimePicker();
            TimePickerPause();
            addSiteDependOnOrderTime();

            Dispatcher.BeginInvoke(new Action(() =>
            {
                _isInitialized = true;
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

        }

        


        private void SetAllDateForThisWeek(object sender, SelectionChangedEventArgs e)
        {
            if(dp_Datum_Mo_Stunden.SelectedDate == null)
            {
                return;
            }
            DateTime DateFirstWeekday = (DateTime)dp_Datum_Mo_Stunden.SelectedDate;

            dp_Datum_Di_Stunden.SelectedDate = DateFirstWeekday.AddDays(1);
            dp_Datum_Mi_Stunden.SelectedDate = DateFirstWeekday.AddDays(2);
            dp_Datum_Do_Stunden.SelectedDate = DateFirstWeekday.AddDays(3);
            dp_Datum_Fr_Stunden.SelectedDate = DateFirstWeekday.AddDays(4);
            dp_Datum_Sa_Stunden.SelectedDate = DateFirstWeekday.AddDays(5);
            dp_Datum_So_Stunden.SelectedDate = DateFirstWeekday.AddDays(6);
        }


        private void SiteSwitched_Stunden(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitialized == false)
            {
                return;
            }

            cb_Siteswitch_Stunden.SelectionChanged -= SiteSwitched_Stunden;

            MainWindow mainWindow = (MainWindow)Application.Current.MainWindow;
            string LastSelectedItem = cb_Siteswitch_Stunden.SelectionBoxItem.ToString();
            string selectedItem = cb_Siteswitch_Stunden.SelectedItem.ToString();
            string selectedItemText = selectedItem.Substring(selectedItem.IndexOf(" ") + 1);
            string ExcelFilePathLoad = "";
            string ExcelFilePathSave = "";

            if (LastSelectedItem == "Woche 1")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 2")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_2.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 3")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_3.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 4")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_4.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 5")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_5.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 6")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_6.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 7")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_7.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }
            else if (LastSelectedItem == "Woche 8")
            {
                ExcelFilePathSave = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_8.xlsm");
                mainWindow.speichern(ExcelFilePathSave, "Stundennachweis");
            }


            if (selectedItemText == "Woche 1")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 2")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_2.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 3")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_3.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 4")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_4.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 5")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_5.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 6")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_6.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 7")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_7.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }
            else if (selectedItemText == "Woche 8")
            {
                ExcelFilePathLoad = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, "Stundennachweis_8.xlsm");
                mainWindow.Laden(ExcelFilePathLoad, "Stundennachweis");
            }



            cb_Siteswitch_Stunden.SelectionChanged += SiteSwitched_Stunden;
        }

        private void addSiteDependOnOrderTime()
        {
            TimeSpan ServiceDurationInDays = GlobalVariables.EndeServiceEinsatz - GlobalVariables.StartServiceEinsatz;

            int Weeks = ServiceDurationInDays.Days / 7;

            for (int i = 0; i < Weeks; i++)
            {
                string quellOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_QuelleVorlagen, "Stundennachweis.xlsm");
                string ZielData = "Stundennachweis.xlsm";
                int x = i + 1;
                if (i != 0) 
                {
                    ZielData = "Stundennachweis_" + x.ToString() + ".xlsm";
                }
                    string zielOrdner = System.IO.Path.Combine(GlobalVariables.Pfad_AuftragsOrdner, ZielData);

                if (!File.Exists(zielOrdner))
                {
                    File.Copy(quellOrdner, zielOrdner);
                }
                
                string item = "cbItem_SiteSwitch_Stunden" + x.ToString();
                
                

                ComboBoxItem Item = (ComboBoxItem)Grid_Stunden.FindName(item);

                Item.Visibility = Visibility.Visible;

            }
        }


        private void TimePicker()
        {
            for (int hour = 0; hour < 24; hour++)
            {
                for (int minute = 0; minute < 60; minute += 15)
                {
                    cb_Anreise_Fahrtbeginn_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Anreise_Fahrtende_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Abreise_Fahrtbeginn_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Abreise_Fahrtende_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Mo_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Di_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Mi_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Do_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Fr_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Sa_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_So_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Mo_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Di_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Mi_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Do_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Fr_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Sa_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_So_Stunden.Items.Add($"{hour:D2}:{minute:D2}"); 
                    cb_Von_Mo_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Di_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Mi_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Do_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Fr_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_Sa_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Von_So_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Mo_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Di_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Mi_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Do_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Fr_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_Sa_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Bis_So_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                }
            }
        }
        private void TimePickerPause()
        {
            for (int hour = 0; hour < 3.2; hour++)
            {
                for (int minute = 0; minute < 60; minute += 15)
                {
                    cb_Anreise_Pause_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Abreise_Pause_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Mo_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Di_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Mi_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Do_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Fr_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Sa_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_So_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Mo_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Di_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Mi_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Do_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Fr_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_Sa_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                    cb_Pause_So_S2_Stunden.Items.Add($"{hour:D2}:{minute:D2}");
                }
            }
        }

        public static TimeSpan[] TäglicheArbeitszeitBerechnen(TimeSpan Arbeitsbeginn, TimeSpan Arbeitsende, TimeSpan Pause, TimeSpan ArbeitsbeginnS2, TimeSpan ArbeitsendeS2, TimeSpan PauseS2)
        {
            TimeSpan[] Zeiten = new TimeSpan[4];
            TimeSpan NormalStd = new TimeSpan(0, 0, 0);
            TimeSpan Ueberstunden = new TimeSpan(0, 0, 0);
            TimeSpan Nachtarbeit = new TimeSpan(0, 0, 0);
            TimeSpan GesamtStd = new TimeSpan(0, 0, 0);

            if (Arbeitsbeginn != TimeSpan.Zero && Arbeitsende != TimeSpan.Zero)
            {
                if (Arbeitsbeginn < GlobalVariables.FruehNacht)
                {
                    Nachtarbeit += (GlobalVariables.FruehNacht - Arbeitsbeginn);
                }
                if (Arbeitsende > GlobalVariables.SpaetNacht)
                {
                    Nachtarbeit += (Arbeitsende - GlobalVariables.SpaetNacht);
                }                
            }
            if (ArbeitsbeginnS2 != TimeSpan.Zero && ArbeitsendeS2 != TimeSpan.Zero)
            {
                if (ArbeitsbeginnS2 < GlobalVariables.FruehNacht)
                {
                    Nachtarbeit += (GlobalVariables.FruehNacht - ArbeitsbeginnS2);
                }
                if (ArbeitsendeS2 > GlobalVariables.SpaetNacht)
                {
                    Nachtarbeit += (ArbeitsendeS2 - GlobalVariables.SpaetNacht);
                }
            }
                GesamtStd = ((Arbeitsende - Arbeitsbeginn) - Pause) + ((ArbeitsendeS2 - ArbeitsbeginnS2) - PauseS2);

            if (GesamtStd > GlobalVariables.RegularStd)
            {
                Ueberstunden += GesamtStd - GlobalVariables.RegularStd;
                NormalStd = GlobalVariables.RegularStd;
            }
            else
            {
                NormalStd = GesamtStd;
            }

            if (NormalStd < GlobalVariables.RegularStd)
            {
                Ueberstunden = NormalStd - GlobalVariables.RegularStd;
            }

            Zeiten[1] = NormalStd;
            Zeiten[0] = Ueberstunden;
            Zeiten[3] = Nachtarbeit;
            Zeiten[2] = GesamtStd;

            return Zeiten;
        }

        public static TimeSpan[] WochendZeitenBerechnen(TimeSpan Arbeitsbeginn, TimeSpan Arbeitsende, TimeSpan Pause, TimeSpan ArbeitsbeginnS2, TimeSpan ArbeitsendeS2, TimeSpan PauseS2)
        {
            TimeSpan[] Zeiten = new TimeSpan[4];
            TimeSpan NormalStd = new TimeSpan(0, 0, 0);
            TimeSpan Nachtarbeit = new TimeSpan(0, 0, 0);
            TimeSpan GesamtStd = new TimeSpan(0, 0, 0);

            if (Arbeitsbeginn != TimeSpan.Zero && Arbeitsende != TimeSpan.Zero)
            {
                if (Arbeitsbeginn < GlobalVariables.FruehNacht)
                {
                    Nachtarbeit += (GlobalVariables.FruehNacht - Arbeitsbeginn);
                }
                if (Arbeitsende > GlobalVariables.SpaetNacht)
                {
                    Nachtarbeit += (Arbeitsende - GlobalVariables.SpaetNacht);
                }
            }
            if (ArbeitsbeginnS2 != TimeSpan.Zero && ArbeitsendeS2 != TimeSpan.Zero)
            {
                if (ArbeitsbeginnS2 < GlobalVariables.FruehNacht)
                {
                    Nachtarbeit += (GlobalVariables.FruehNacht - ArbeitsbeginnS2);
                }
                if (ArbeitsendeS2 > GlobalVariables.SpaetNacht)
                {
                    Nachtarbeit += (ArbeitsendeS2 - GlobalVariables.SpaetNacht);
                }
            }

            GesamtStd = ((Arbeitsende - Arbeitsbeginn) - Pause) + ((ArbeitsendeS2 - ArbeitsbeginnS2) - PauseS2);

            Zeiten[0] = GesamtStd;
            Zeiten[1] = NormalStd;
            Zeiten[2] = Nachtarbeit;

            return Zeiten;
        }

        private void cb_Anreise_Pause_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Gesamt_Anreisedauer;
            TimeSpan An_Fahrtbeginn;
            TimeSpan.TryParse(cb_Anreise_Fahrtbeginn_Stunden.Text, out An_Fahrtbeginn);
            TimeSpan An_Fahrtende;
            TimeSpan.TryParse(cb_Anreise_Fahrtende_Stunden.Text, out An_Fahrtende);
            TimeSpan An_Pause;
            string temp = cb_Anreise_Pause_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out An_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24, 0, 0);
            if (dp_AnreiseDatum_Stunden.Text != dp_AnreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Anreisedauer = (TempAnderesDatum - An_Fahrtbeginn - An_Pause) + An_Fahrtende;
            }
            else
            {
                Gesamt_Anreisedauer = (An_Fahrtende - An_Fahrtbeginn) - An_Pause;
            }
            tb_Anreisedauer_Gesamt_Stunden.Text = Gesamt_Anreisedauer.ToString();

        }

        private void cb_Anreise_Fahrtende_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Gesamt_Anreisedauer;
            TimeSpan An_Fahrtbeginn;
            TimeSpan.TryParse(cb_Anreise_Fahrtbeginn_Stunden.Text, out An_Fahrtbeginn);
            TimeSpan An_Fahrtende;
            string temp = cb_Anreise_Fahrtende_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out An_Fahrtende);
            TimeSpan An_Pause;
            TimeSpan.TryParse(cb_Anreise_Pause_Stunden.Text, out An_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24, 0, 0);
            if (dp_AnreiseDatum_Stunden.Text != dp_AnreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Anreisedauer = (TempAnderesDatum - An_Fahrtbeginn - An_Pause) + An_Fahrtende;
            }
            else
            {
                Gesamt_Anreisedauer = (An_Fahrtende - An_Fahrtbeginn) - An_Pause;
            }
            tb_Anreisedauer_Gesamt_Stunden.Text = Gesamt_Anreisedauer.ToString();
        }

        private void cb_Anreise_Fahrtbeginn_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Gesamt_Anreisedauer;
            TimeSpan An_Fahrtbeginn;
            string temp = cb_Anreise_Fahrtbeginn_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out An_Fahrtbeginn);
            TimeSpan An_Fahrtende;
            TimeSpan.TryParse(cb_Anreise_Fahrtende_Stunden.Text, out An_Fahrtende);
            TimeSpan An_Pause;
            TimeSpan.TryParse(cb_Abreise_Pause_Stunden.Text, out An_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24, 0, 0);
            if (dp_AnreiseDatum_Stunden.Text != dp_AnreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Anreisedauer = (TempAnderesDatum - An_Fahrtbeginn - An_Pause) + An_Fahrtende;
            }
            else
            {
                Gesamt_Anreisedauer = (An_Fahrtende - An_Fahrtbeginn) - An_Pause;
            }
            tb_Anreisedauer_Gesamt_Stunden.Text = Gesamt_Anreisedauer.ToString();
        }
        private void cb_Abreise_Pause_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            TimeSpan Gesamt_Abreisedauer;
            TimeSpan Ab_Fahrtbeginn;
            TimeSpan.TryParse(cb_Abreise_Fahrtbeginn_Stunden.Text, out Ab_Fahrtbeginn);
            TimeSpan Ab_Fahrtende;
            TimeSpan.TryParse(cb_Abreise_Fahrtende_Stunden.Text, out Ab_Fahrtende);
            TimeSpan Ab_Pause;
            string temp = cb_Abreise_Pause_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Ab_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24,0,0);
            if (dp_AbreiseDatum_Stunden.Text != dp_AbreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Abreisedauer = (TempAnderesDatum - Ab_Fahrtbeginn - Ab_Pause) + Ab_Fahrtende;
            }
            else
            {
                Gesamt_Abreisedauer = (Ab_Fahrtende - Ab_Fahrtbeginn) - Ab_Pause;
            }
            tb_Abreisedauer_Gesamt_Stunden.Text = Gesamt_Abreisedauer.ToString();
        }

        private void cb_Abreise_Fahrtende_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Gesamt_Abreisedauer;
            TimeSpan Ab_Fahrtbeginn;
            TimeSpan.TryParse(cb_Abreise_Fahrtbeginn_Stunden.Text, out Ab_Fahrtbeginn);
            TimeSpan Ab_Fahrtende;
            string temp = cb_Abreise_Fahrtende_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Ab_Fahrtende);
            TimeSpan Ab_Pause;
            TimeSpan.TryParse(cb_Abreise_Pause_Stunden.Text, out Ab_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24,0,0);
            if (dp_AbreiseDatum_Stunden.Text != dp_AbreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Abreisedauer = (TempAnderesDatum - Ab_Fahrtbeginn - Ab_Pause) + Ab_Fahrtende;
            }
            else
            {
                Gesamt_Abreisedauer = (Ab_Fahrtende - Ab_Fahrtbeginn) - Ab_Pause;
            }
            tb_Abreisedauer_Gesamt_Stunden.Text = Gesamt_Abreisedauer.ToString();
        }

        private void cb_Abreise_Fahrtbeginn_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Gesamt_Abreisedauer;
            TimeSpan Ab_Fahrtbeginn;
            string temp = cb_Abreise_Fahrtbeginn_Stunden.SelectedItem as string; //um das neu ausgewählte Item zu nutzen
            TimeSpan.TryParse(temp, out Ab_Fahrtbeginn);
            TimeSpan Ab_Fahrtende;
            TimeSpan.TryParse(cb_Abreise_Fahrtende_Stunden.Text, out Ab_Fahrtende);
            TimeSpan Ab_Pause;
            TimeSpan.TryParse(cb_Abreise_Pause_Stunden.Text, out Ab_Pause);
            TimeSpan TempAnderesDatum = new TimeSpan(24,0,0);
            if (dp_AbreiseDatum_Stunden.Text != dp_AbreiseDatumAnkunft_Stunden.Text)
            {
                Gesamt_Abreisedauer = (TempAnderesDatum - Ab_Fahrtbeginn - Ab_Pause) + Ab_Fahrtende;
            }
            else
            {
                Gesamt_Abreisedauer = (Ab_Fahrtende - Ab_Fahrtbeginn) - Ab_Pause;
            }
            tb_Abreisedauer_Gesamt_Stunden.Text = Gesamt_Abreisedauer.ToString();
        }
        
        //EventHandler Montag
        private void cb_Von_Mo_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;
            
            string temp = cb_Von_Mo_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Mo_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mo_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Mo_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Mo_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mo_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Mo_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Von_Mo_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            
            TimeSpan.TryParse(cb_Von_Mo_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            string temp = cb_Von_Mo_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Mo_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mo_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Mo_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp , out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Mo_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mo_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Mo_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mo_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mo_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mo_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mo_Stunden.Text = Zeiten[3].ToString();
        }


        //EventHandler Dienstag
        private void cb_Von_Di_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_Di_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Di_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Di_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Di_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Di_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Di_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Di_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Di_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Di_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Di_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Di_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Di_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Di_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Di_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Di_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Di_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Di_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Di_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Di_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Di_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Von_Di_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;


            TimeSpan.TryParse(cb_Von_Di_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Di_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Di_Stunden.Text, out Pause);
            string temp = cb_Von_Di_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Di_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Di_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Di_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Di_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Di_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Di_Stunden.Text, out Pause);

            TimeSpan.TryParse(cb_Von_Di_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Di_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Di_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Di_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Di_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Di_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Di_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Di_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Di_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Di_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Di_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Di_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Di_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Di_Stunden.Text = Zeiten[3].ToString();
        }



        //EventHandler Mittwoch
        private void cb_Von_Mi_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_Mi_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mi_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mi_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mi_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mi_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mi_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Mi_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mi_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Mi_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mi_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mi_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mi_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mi_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Mi_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mi_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mi_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Mi_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Mi_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mi_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mi_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Von_Mi_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;


            TimeSpan.TryParse(cb_Von_Mi_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mi_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mi_Stunden.Text, out Pause);
            string temp = cb_Von_Mi_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mi_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mi_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Mi_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mi_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mi_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mi_Stunden.Text, out Pause);

            TimeSpan.TryParse(cb_Von_Mi_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Mi_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mi_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Mi_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Mi_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mi_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mi_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mi_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mi_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Mi_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Mi_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Mi_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Mi_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Mi_Stunden.Text = Zeiten[3].ToString();
        }


        //EventHandler Donnerstag
        private void cb_Von_Do_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_Do_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Do_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Do_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Do_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Do_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Do_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Do_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Do_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Do_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Do_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Do_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Do_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Do_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Do_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Do_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Do_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Do_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Do_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Do_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Do_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Von_Do_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;


            TimeSpan.TryParse(cb_Von_Do_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Do_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Do_Stunden.Text, out Pause);
            string temp = cb_Von_Do_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Do_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Do_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Do_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Do_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Do_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Do_Stunden.Text, out Pause);

            TimeSpan.TryParse(cb_Von_Do_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Do_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Do_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Do_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Do_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Do_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Do_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Do_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Do_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Do_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Do_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Do_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Do_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Do_Stunden.Text = Zeiten[3].ToString();
        }



        //EventHandler Freitag
        private void cb_Von_Fr_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_Fr_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Fr_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Fr_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Fr_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Fr_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Fr_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Fr_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Fr_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Fr_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Fr_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Fr_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Fr_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Fr_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Fr_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Fr_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Fr_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Fr_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Fr_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Fr_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Fr_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Von_Fr_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;


            TimeSpan.TryParse(cb_Von_Fr_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Fr_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Fr_Stunden.Text, out Pause);
            string temp = cb_Von_Fr_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Fr_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Fr_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Bis_Fr_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Fr_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Fr_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Fr_Stunden.Text, out Pause);

            TimeSpan.TryParse(cb_Von_Fr_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Fr_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Fr_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }

        private void cb_Pause_Fr_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Fr_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Fr_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Fr_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Fr_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Fr_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Fr_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Fr_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Fr_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Fr_Stunden.Text = Zeiten[2].ToString();
            tb_Nachtarbeit_Fr_Stunden.Text = Zeiten[3].ToString();
        }


        //EventHandler Samstag
        private void cb_Von_Sa_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_Sa_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Sa_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Sa_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Sa_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Sa_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Sa_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Bis_Sa_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Sa_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_Sa_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Sa_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Sa_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Sa_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Sa_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Pause_Sa_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Sa_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Sa_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_Sa_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_Sa_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Sa_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Sa_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Von_Sa_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            
            TimeSpan.TryParse(cb_Von_Sa_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Sa_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Sa_Stunden.Text, out Pause);
            string temp = cb_Von_Sa_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Sa_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Sa_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Bis_Sa_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Sa_Stunden.Text, out Arbeitsbeginn);            
            TimeSpan.TryParse(cb_Bis_Sa_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Sa_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Sa_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_Sa_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Sa_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Pause_Sa_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_Sa_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Sa_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Sa_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Sa_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Sa_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_Sa_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_Sa_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_Sa_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_Sa_Stunden.Text = Zeiten[2].ToString();
        }


        //EventHandler Sonntag
        private void cb_Von_So_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            string temp = cb_Von_So_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_So_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_So_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_So_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_So_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_So_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Bis_So_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_So_Stunden.Text, out Arbeitsbeginn);
            string temp = cb_Bis_So_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_So_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_So_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_So_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_So_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Pause_So_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_So_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_So_Stunden.Text, out Arbeitsende);
            string temp = cb_Pause_So_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Pause);
            TimeSpan.TryParse(cb_Von_So_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_So_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_So_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Von_So_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;


            TimeSpan.TryParse(cb_Von_So_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_So_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_So_Stunden.Text, out Pause);
            string temp = cb_Von_So_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_So_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_So_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Bis_So_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_So_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_So_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_So_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_So_S2_Stunden.Text, out ArbeitsbeginnS2);
            string temp = cb_Bis_So_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_So_S2_Stunden.Text, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }

        private void cb_Pause_So_S2_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;

            TimeSpan.TryParse(cb_Von_So_Stunden.Text, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_So_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_So_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_So_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_So_S2_Stunden.Text, out ArbeitsendeS2);
            string temp = cb_Pause_So_S2_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out PauseS2);

            TimeSpan[] Zeiten = WochendZeitenBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            tb_Ueberstunden_So_Stunden.Text = Zeiten[0].ToString();
            tb_NormalStd_So_Stunden.Text = Zeiten[1].ToString();
            tb_GesamtStd_So_Stunden.Text = Zeiten[0].ToString();
            tb_Nachtarbeit_So_Stunden.Text = Zeiten[2].ToString();
        }
        //EventHandler Tage Vorbei
        

        private void GesamtRegStd()
        {
            TimeSpan Mo;
            TimeSpan Di;
            TimeSpan Mi;
            TimeSpan Do;
            TimeSpan Fr;
            TimeSpan Gesamt;

            TimeSpan.TryParse(tb_NormalStd_Mo_Stunden.Text, out Mo);
            TimeSpan.TryParse(tb_NormalStd_Di_Stunden.Text, out Di);
            TimeSpan.TryParse(tb_NormalStd_Mi_Stunden.Text, out Mi);
            TimeSpan.TryParse(tb_NormalStd_Do_Stunden.Text, out Do);
            TimeSpan.TryParse(tb_NormalStd_Fr_Stunden.Text, out Fr);

            Gesamt = Mo + Di + Mi + Do + Fr;

            tb_GesamteWoche_NormalStd_Stunden.Text = Gesamt.ToString();
        }

        private void tb_NormalStd_Mo_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtRegStd();
        }

        private void tb_NormalStd_Di_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtRegStd();
        }

        private void tb_NormalStd_Mi_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtRegStd();
        }

        private void tb_NormalStd_Do_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtRegStd();
        }

        private void tb_NormalStd_Fr_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtRegStd();
        }

        private void GesamtUeberStd()
        {
            TimeSpan Mo;
            TimeSpan Di;
            TimeSpan Mi;
            TimeSpan Do;
            TimeSpan Fr;
            TimeSpan Sa;
            TimeSpan So;
            TimeSpan Gesamt;
           

            TimeSpan.TryParse(tb_Ueberstunden_Mo_Stunden.Text, out Mo);
            TimeSpan.TryParse(tb_Ueberstunden_Di_Stunden.Text, out Di);
            TimeSpan.TryParse(tb_Ueberstunden_Mi_Stunden.Text, out Mi);
            TimeSpan.TryParse(tb_Ueberstunden_Do_Stunden.Text, out Do);
            TimeSpan.TryParse(tb_Ueberstunden_Fr_Stunden.Text, out Fr);
            TimeSpan.TryParse(tb_Ueberstunden_Sa_Stunden.Text, out Sa);
            TimeSpan.TryParse(tb_Ueberstunden_So_Stunden.Text, out So);

            Gesamt = Mo + Di + Mi + Do + Fr + Sa + So;

            tb_GesamteWoche_UeberStd_Stunden.Text = Gesamt.ToString();
        }

        private void tb_Ueberstunden_Mo_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_Di_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_Mi_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_Do_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_Fr_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_Sa_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void tb_Ueberstunden_So_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtUeberStd();
        }

        private void GesamtNachtStd()
        {
            TimeSpan Mo;
            TimeSpan Di;
            TimeSpan Mi;
            TimeSpan Do;
            TimeSpan Fr;
            TimeSpan Sa;
            TimeSpan So;
            TimeSpan Gesamt;

            TimeSpan.TryParse(tb_Nachtarbeit_Mo_Stunden.Text, out Mo);
            TimeSpan.TryParse(tb_Nachtarbeit_Di_Stunden.Text, out Di);
            TimeSpan.TryParse(tb_Nachtarbeit_Mi_Stunden.Text, out Mi);
            TimeSpan.TryParse(tb_Nachtarbeit_Do_Stunden.Text, out Do);
            TimeSpan.TryParse(tb_Nachtarbeit_Fr_Stunden.Text, out Fr);
            TimeSpan.TryParse(tb_Nachtarbeit_Sa_Stunden.Text, out Sa);
            TimeSpan.TryParse(tb_Nachtarbeit_So_Stunden.Text, out So);

            Gesamt = Mo + Di + Mi + Do + Fr + Sa + So;

            tb_GesamteWoche_NachtStd_Stunden.Text = Gesamt.ToString();
        }

        private void tb_Nachtarbeit_Mo_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_Di_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_Mi_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_Do_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_Fr_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_Sa_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void tb_Nachtarbeit_So_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamtNachtStd();
        }

        private void GesamteStd()
        {
            TimeSpan Mo;
            TimeSpan Di;
            TimeSpan Mi;
            TimeSpan Do;
            TimeSpan Fr;
            TimeSpan Sa;
            TimeSpan So;
            TimeSpan Gesamt;

            TimeSpan.TryParse(tb_GesamtStd_Mo_Stunden.Text, out Mo);
            TimeSpan.TryParse(tb_GesamtStd_Di_Stunden.Text, out Di);
            TimeSpan.TryParse(tb_GesamtStd_Mi_Stunden.Text, out Mi);
            TimeSpan.TryParse(tb_GesamtStd_Do_Stunden.Text, out Do);
            TimeSpan.TryParse(tb_GesamtStd_Fr_Stunden.Text, out Fr);
            TimeSpan.TryParse(tb_GesamtStd_Sa_Stunden.Text, out Sa);
            TimeSpan.TryParse(tb_GesamtStd_So_Stunden.Text, out So);

            Gesamt = Mo + Di + Mi + Do + Fr + Sa + So;

            tb_GesamteWoche_AlleStd_Stunden.Text = Gesamt.ToString();
        }

        private void tb_GesamtStd_Mo_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_Di_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_Mi_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_Do_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_Fr_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_Sa_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void tb_GesamtStd_So_Stunden_TextChanged(object sender, TextChangedEventArgs e)
        {
            GesamteStd();
        }

        private void Schicht2Hinzufügen_Mo(object sender, RoutedEventArgs e)
        {
            cb_Von_Mo_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Mo_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Mo_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Mo_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Mo.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_Di(object sender, RoutedEventArgs e)
        {
            cb_Von_Di_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Di_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Di_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Di_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Di.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_Mi(object sender, RoutedEventArgs e)
        {
            cb_Von_Mi_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Mi_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Mi_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Mi_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Mi.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_Do(object sender, RoutedEventArgs e)
        {
            cb_Von_Do_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Do_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Do_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Do_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Do.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_Fr(object sender, RoutedEventArgs e)
        {
            cb_Von_Fr_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Fr_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Fr_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Fr_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Fr.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_Sa(object sender, RoutedEventArgs e)
        {
            cb_Von_Sa_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_Sa_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_Sa_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_Sa_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_Sa.Visibility = Visibility.Visible;
        }
        private void Schicht2Hinzufügen_So(object sender, RoutedEventArgs e)
        {
            cb_Von_So_S2_Stunden.Visibility = Visibility.Visible;
            cb_Bis_So_S2_Stunden.Visibility = Visibility.Visible;
            cb_Pause_So_S2_Stunden.Visibility = Visibility.Visible;
            tb_Bemerkung_So_S2_Stunden.Visibility = Visibility.Visible;
            lbl_Schicht2_So.Visibility = Visibility.Visible;
        }
    }
}
