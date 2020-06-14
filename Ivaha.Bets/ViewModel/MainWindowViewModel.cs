using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml;
using System.Xml.Serialization;
using Bluegrams.Application;
using Ivaha.Bets.Model;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Ivaha.Bets.ViewModel
{
    public  class   MainWindowViewModel : ViewModelBase
    {
        public  class   ComandNameGroup : ViewModelBase
        {
            const   byte            _MAX_CHARS_IN_GROUP_NAME        =   40;
            const   string          _NAMES_DELIMITER                =   ",";

            public  string          GroupName       { get; set; }
            public  string          GroupValues     { get; set; }
            public  List<string>    Names           { get; set; }
            public  bool            IsReadOnly      { get; set; }   =   true;
            public  Brush           Background      { get; set; }
            public  bool            Focused         { get; set; }
            public  double          TextBoxWidth    { get; set; }
            public  Visibility      EditVisibility  { get; set; }
            public  Visibility      SaveVisibility  { get; set; }   =   Visibility.Collapsed;

            public                  ComandNameGroup () : base (Application.Current?.MainWindow, CommandOwnerType)
            {
                RegCommand(EditGroup        , editGroup         );
                RegCommand(SaveGroup        , saveGroup         );
                RegCommand(RemoveGroup      , removeGroup       );

                PropertyChanged    +=   (s,e) =>
                {
                    switch (e.PropertyName)
                    {
                        case nameof(IsReadOnly):
                            Background      =   IsReadOnly ? Brushes.Transparent : SystemColors.WindowBrush;
                            EditVisibility  =   IsReadOnly ? Visibility.Visible : Visibility.Collapsed;
                            SaveVisibility  =  !IsReadOnly ? Visibility.Visible : Visibility.Collapsed;

                            if (!IsReadOnly)
                            {
                                GroupName   =   GroupValues;
                                Focused     =   false;
                                Focused     =   true;
                            }
                            else
                            {
                                GroupValues =   GroupName;
                                GroupName   =   Truncate(GroupValues, _MAX_CHARS_IN_GROUP_NAME);
                            }
                            break;

                        case nameof(GroupValues):
                            Names           =   GroupValues.Split(new []{_NAMES_DELIMITER}, StringSplitOptions.RemoveEmptyEntries).Select(str => str.Trim()).ToList();
                            GroupName       =   Truncate(GroupValues, _MAX_CHARS_IN_GROUP_NAME);

                            MainWindowViewModel._VM.SaveSameNames();
                            break;
                    }
                };
            }

            static  Type            CommandOwnerType    =   typeof(ComandNameGroup);

            public RoutedUICommand  EditGroup           { get; }    =   new RoutedUICommand("Редактировать группу"  , "EditGroup"         , CommandOwnerType);
            public RoutedUICommand  SaveGroup           { get; }    =   new RoutedUICommand("Сохранить изменения"   , "SaveGroup"         , CommandOwnerType);
            public RoutedUICommand  RemoveGroup         { get; }    =   new RoutedUICommand("Удалить группу"        , "RemoveGroup"       , CommandOwnerType);

            private void            editGroup           (object sender, ExecutedRoutedEventArgs e)  => IsReadOnly = false;
            private void            saveGroup           (object sender, ExecutedRoutedEventArgs e)  => IsReadOnly = true;
            private void            removeGroup         (object sender, ExecutedRoutedEventArgs e)
            {
                if (MessageBox.Show($"Вы действительно хотите убрать группу одноименных команд ({string.Join(", ", Names)})?", 
                    MainControl?.Title, MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                    return;

                _VM.Groups.Remove(this);
            }

            private static  string  Truncate            (string value, int maxChars)                =>  value.Length <= maxChars ? value : value.Substring(0, maxChars) + "...";
        }

        public  static  MainWindowViewModel _VM;

        public  string          SourceFileName      { get; set; }   //=   "source.xlsx";
        public  string          ResultFileName      { get; set; }   //=   "result.xlsx";
        public  string          Logs                { get; set; }   =   string.Empty;
        public  ObservableCollection<ComandNameGroup>
                                Groups              { get; set; }   =   new ObservableCollection<ComandNameGroup>();

        private Dictionary<string, Team>    Teams   =   new Dictionary<string, Team>();
        private bool                        Loaded  =   false;

        public                  MainWindowViewModel () : base (Application.Current?.MainWindow, CommandOwnerType)
        {
            _VM =   this;

            RegCommand(OpenFileDialog   , openFileDialog    );
            RegCommand(SaveFileDialog   , saveFileDialog    );

            //PortableSettingsProvider.ApplyProvider(Properties.Settings.Default);
            //Properties.Settings.Default.Reset();
            ReadSameNames();

            PropertyChanged    +=   (s,e) =>
            {
                switch (e.PropertyName)
                {
                    case nameof(Groups):
                        SaveSameNames();
                        break;
                }
            };

            MainControl.Loaded +=   (s,e) => Loaded = true;

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

        public  void            ReadSameNames       ()
        {
            foreach (var names in Properties.Settings.Default.SameNames.Split(new []{ Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                Groups.Add(new ComandNameGroup(){ GroupValues = names });
        }
        public  void            SaveSameNames       ()
        {
            if (!Loaded)
                return;

            try
            {
                Properties.Settings.Default.SameNames   =   string.Join(Environment.NewLine, Groups.Select(g => string.Join(", ", g.Names)));

                if (!Properties.Settings.Default.Context.IsReadOnly)
                    Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }
        private MessageBoxResult ShowMessageInvoke  (Window owner, string message, MessageBoxButton buttons = MessageBoxButton.OK)  =>  
            Application.Current?.Dispatcher.Invoke(() => owner == null 
                                                       ? MessageBox.Show(message, MainControl?.Title, buttons) 
                                                       : MessageBox.Show(owner, message, MainControl?.Title, buttons)) 
            ?? MessageBoxResult.OK;

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

            var cts         =   new CancellationTokenSource();
            var token       =   cts.Token;
            var caption     =   MainControl?.Title;

            ProgressWindow.Run((aProgress, aCts, aToken, aWindow) =>
            {
                if (Excel.ExportToExcel(Teams.Select(kvp => kvp.Value).ToArray(), ResultFileName, caption, onLog, aToken))
                    ShowMessageInvoke(aWindow, "Файл успешно сохранен");
            }, cts, token,
            (ex, win) => ShowMessageInvoke(win, $"Ошибка при сохранении файла:{Environment.NewLine}{ex.Message}"));
        }
        private void            onLog               (string message)
        {
            Logs   +=  $"{(Logs.Length > 0 ? Environment.NewLine : "")}{message}";
        }

        #endregion
    }
}
