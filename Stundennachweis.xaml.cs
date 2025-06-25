using Microsoft.Office.Interop.Excel;
using Microsoft.Win32.SafeHandles;
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
    public partial class Stundennachweis : UserControl
    {
        private bool _isInitialized = false;
        public Stundennachweis()
        {
            InitializeComponent();
            TimePicker();//Add all the TimePicker Items to the ComboBoxes for entering the Working hours
            TimePickerPause();//add all the TimePicker Items to the Pause ComboBoxes
            addSiteDependOnOrderTime();//Depending on how many weeks the service order takes, add the corresponding number of sites to the ComboBox for switching between the weeks

            Dispatcher.BeginInvoke(new System.Action(() =>
            {
                _isInitialized = true;
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

        }

        private void SetAllDateForThisWeek(object sender, SelectionChangedEventArgs e)
        {//if the Date of the first weekday (Monday) is selected, set all other dates of the week accordingly
            if (dp_Datum_Mo_Stunden.SelectedDate == null)
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
        {//this method is called when the user switches between the weeks in the ComboBox
            if (_isInitialized == false)
            {
                return;
            }
            //Deactivate the SelectionChanged event to prevent a Loop
            cb_Siteswitch_Stunden.SelectionChanged -= SiteSwitched_Stunden;

            MainWindow mainWindow = (MainWindow)System.Windows.Application.Current.MainWindow;
            string LastSelectedItem = cb_Siteswitch_Stunden.SelectionBoxItem.ToString();
            string selectedItem = cb_Siteswitch_Stunden.SelectedItem.ToString();
            string selectedItemText = selectedItem.Substring(selectedItem.IndexOf(" ") + 1);
            string ExcelFilePathLoad = "";
            string ExcelFilePathSave = "";
            //Set the Path for the Excel file to save the data depending on the last selected item in the ComboBox
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

            //Set the Path for the Excel file to load the data depending on the selected item in the ComboBox
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

            //Set all the Information from the GlobalVariables to the TextBoxes and ComboBoxes in the UserControl
            tb_Servicetechiker_Stunden.Text = GlobalVariables.ServiceTechnicker;
            tb_Servicetechiker_Stunden.Focusable = false;
            tb_Kunde_Stunden.Text = GlobalVariables.Kunde;
            tb_Kunde_Stunden.Focusable = false;
            tb_Ansprechpartner_Stunden.Text = GlobalVariables.Ansprechpartner;
            tb_Ansprechpartner_Stunden.Focusable = false;
            tb_Anschrift_1_Stunden.Text = GlobalVariables.Anschrift_1;
            tb_Anschrift_1_Stunden.Focusable = false;
            tb_Anschrift_2_Stunden.Text = GlobalVariables.Anschrift_2;
            tb_Anschrift_2_Stunden.Focusable = false;
            if (GlobalVariables.Anreise != "")
            {
                cb_Verkehrsmittel_Stunden.Text = GlobalVariables.Anreise;
                cb_Verkehrsmittel_Stunden.Focusable = false;
            }

            //Reactivate the SelectionChanged event after the data has been loaded
            cb_Siteswitch_Stunden.SelectionChanged += SiteSwitched_Stunden;
        }

        private void addSiteDependOnOrderTime()
        {
            //Calculate the number of weeks and safe it as a whole number in the variable Weeks
            TimeSpan ServiceDurationInDays = GlobalVariables.EndeServiceEinsatz - GlobalVariables.StartServiceEinsatz;
            double weeksnotRounded = ServiceDurationInDays.TotalDays/7;
            int Weeks = (int)Math.Ceiling(weeksnotRounded);

            for (int i = 0; i < Weeks; i++)
            {
                //Set the Dataname depending on the number of weeks
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

                //Make the Item Visible depending on the number of weeks
                string item = "cbItem_SiteSwitch_Stunden" + x.ToString();                

                ComboBoxItem Item = (ComboBoxItem)Grid_Stunden.FindName(item);

                Item.Visibility = Visibility.Visible;

            }
        }

        private void TimePicker()
        {//Function to add all the TimePicker Items to the ComboBoxes for entering the Working hours
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
        {//Function to add all the TimePicker Items to the Pause ComboBoxes
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
            //Calculate all the Working hours, Normal hours, Overtime and Night work
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
            //Safe the calculated values in the array
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
            // Calculate all the Working hours, Normal hours, Overtime and Night work
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
        //EventHandler for Arrival and Departure to Calculate the total travel time every time a ComboBox selection changes
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
        //***** End of the EventHandler for Arrival and Departure *****

        //Eventhandler for the Working hours of all week days Only one is commented because the only difference is that the Changed TimeSpan is different and safed diferently

        //EventHandler Montag
        private void cb_Von_Mo_Stunden_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TimeSpan Arbeitsbeginn;
            TimeSpan Arbeitsende;
            TimeSpan Pause;
            TimeSpan ArbeitsbeginnS2;
            TimeSpan ArbeitsendeS2;
            TimeSpan PauseS2;
            //Safe all the Information out of the UserControl to Calculate the Working hours
            string temp = cb_Von_Mo_Stunden.SelectedItem as string;
            TimeSpan.TryParse(temp, out Arbeitsbeginn);
            TimeSpan.TryParse(cb_Bis_Mo_Stunden.Text, out Arbeitsende);
            TimeSpan.TryParse(cb_Pause_Mo_Stunden.Text, out Pause);
            TimeSpan.TryParse(cb_Von_Mo_S2_Stunden.Text, out ArbeitsbeginnS2);
            TimeSpan.TryParse(cb_Bis_Mo_S2_Stunden.Text, out ArbeitsendeS2);
            TimeSpan.TryParse(cb_Pause_Mo_S2_Stunden.Text, out PauseS2);

            //Call the Function to Calculate the Working hours
            TimeSpan[] Zeiten = TäglicheArbeitszeitBerechnen(Arbeitsbeginn, Arbeitsende, Pause, ArbeitsbeginnS2, ArbeitsendeS2, PauseS2);

            //Pase the calculated values into the TextBoxes
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
        //End of the Eventhandler for all days

        private void GesamtRegStd()
        {//Add all normal working hours of the week to a Sum
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
        // EventHandler to Calculate the total normal working hours for the week
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
        {//Funktion to Calculate the total overtime hours for the week
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
        //EventHandler to Calculate the total overtime hours for the week
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
        {//Funktion to Calculate the total night work hours for the week
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
        //EventHandler to Calculate the total night work hours for the week
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
        {//Funktion to Calculate the total working hours for the week
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
        //EventHandler to Calculate the total working hours for the week
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

        //EventHandler to make the second shift visible for each day
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
