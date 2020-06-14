using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using Ivaha.Bets.Model;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Ivaha.Bets.ViewModel
{
    public  class   MainWindowViewModel : ViewModelBase
    {
        public  string          SourceFileName      { get; set; }   //=   "source.xlsx";
        public  string          ResultFileName      { get; set; }   //=   "result.xlsx";
        public  string          Logs                { get; set; }   =   string.Empty;

        private Dictionary<string, Team>    Teams   =   new Dictionary<string, Team>();

        public                  MainWindowViewModel () : base (Application.Current?.MainWindow, CommandOwnerType)
        {
            RegCommand(OpenFileDialog   , openFileDialog    );
            RegCommand(SaveFileDialog   , saveFileDialog    );

            ////////////////////
        ////Teams.Add("BLACKSTAR98"                                 , new Team("BLACKSTAR98"));
        ////Teams.Add("FLEWLESS_PHOENIX"                            , new Team("FLEWLESS_PHOENIX"));
        ////Teams.Add("CHELLOVEKK"                                  , new Team("CHELLOVEKK"));
        ////Teams.Add("FEARGGWP"                                    , new Team("FEARGGWP"));
        ////Teams.Add("&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;", new Team("&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"));
        ////Teams.Add("TAKA"                                        , new Team("TAKA"));

        ////Teams["BLACKSTAR98"].Winners        =   new List<Team>(){ Teams["FLEWLESS_PHOENIX"], Teams["FEARGGWP"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"], Teams["CHELLOVEKK"] };
        ////Teams["BLACKSTAR98"].Losers         =   new List<Team>(){ Teams["FLEWLESS_PHOENIX"], Teams["FEARGGWP"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"], Teams["CHELLOVEKK"] };

        ////Teams["FLEWLESS_PHOENIX"].Winners   =   new List<Team>(){ Teams["BLACKSTAR98"], Teams["FEARGGWP"], Teams["TAKA"] };
        ////Teams["FLEWLESS_PHOENIX"].Losers    =   new List<Team>(){ Teams["BLACKSTAR98"], Teams["FEARGGWP"], Teams["TAKA"], Teams["CHELLOVEKK"] };
        ////Teams["FLEWLESS_PHOENIX"].Tied      =   new List<Team>(){ Teams["CHELLOVEKK"] };

        ////Teams["CHELLOVEKK"].Winners         =   new List<Team>(){ Teams["FEARGGWP"], Teams["BLACKSTAR98"], Teams["FLEWLESS_PHOENIX"] };
        ////Teams["CHELLOVEKK"].Losers          =   new List<Team>(){ Teams["FEARGGWP"], Teams["BLACKSTAR98"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"] };
        ////Teams["CHELLOVEKK"].Tied            =   new List<Team>(){ Teams["TAKA"], Teams["FLEWLESS_PHOENIX"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"] };

        ////foreach (var t in Teams.Values)
        ////    t.MakeAllLists();
            ////////////////////
        }

        #region Commands    |

        static  Type            CommandOwnerType    =   typeof(MainWindowViewModel);

        public RoutedUICommand  OpenFileDialog      { get; }    =   new RoutedUICommand("", "OpenFileDialog"    , CommandOwnerType);
        public RoutedUICommand  SaveFileDialog      { get; }    =   new RoutedUICommand("", "SaveFileDialog"    , CommandOwnerType);

        private void            openFileDialog      (object sender, ExecutedRoutedEventArgs e)
        {
            var dlg         =   new OpenFileDialog(){ Filter = "Excel files (*.xlsx)|*.xlsx" };

            if (dlg.ShowDialog() != true || string.IsNullOrEmpty(dlg.FileName) || !File.Exists(dlg.FileName))
                return;

            SourceFileName  =   dlg.FileName;

            // Try load xls
            try
            {
                Excel.ImportFromExcel(out var teams, SourceFileName, MainControl?.Title, onLog);
                Teams   =   teams.ToDictionary(t => t.Name, t => t);
            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Ошибка при загрузке файла:\n{ex.Message}", MainControl?.Title);
                Log.Error(ex);
            }
        }
        private void            saveFileDialog      (object sender, ExecutedRoutedEventArgs e)
        {
            var dlg         =   new SaveFileDialog(){ Filter = "Excel files (*.xlsx)|*.xlsx" };

            if (dlg.ShowDialog() != true || string.IsNullOrEmpty(dlg.FileName))
                return;

            ResultFileName  =   dlg.FileName;

            Excel.ExportToExcel(Teams.Select(kvp => kvp.Value).ToArray(), ResultFileName, MainControl?.Title, onLog);
        }
        private void            onLog               (string message)
        {
            Logs   +=  $"{(Logs.Length > 0 ? Environment.NewLine : "")}{message}";
        }

        #endregion
    }
}
