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
        readonly    Color       _TEAM_COLOR1        =   Color.FromArgb(251, 228, 213);
        readonly    Color       _TEAM_COLOR2        =   Color.FromArgb(226, 239, 217);
        readonly    Color       _WINERSLOSERS_COLOR =   Color.FromArgb(248, 203, 172);
        readonly    Color       _WINERS_COLOR       =   Color.FromArgb(197, 224, 178);

        public  string          SourceFileName      { get; set; }
        public  string          ResultFileName      { get; set; }

        private Dictionary<string, Team>    Teams   =   new Dictionary<string, Team>();

        public                  MainWindowViewModel () : base (Application.Current?.MainWindow, CommandOwnerType)
        {
            RegCommand(OpenFileDialog   , openFileDialog    );
            RegCommand(SaveFileDialog   , saveFileDialog    );

            ////////////////////
            ResultFileName  =   "result.xlsx";

            Teams.Add("BLACKSTAR98"                                 , new Team("BLACKSTAR98"));
            Teams.Add("FLEWLESS_PHOENIX"                            , new Team("FLEWLESS_PHOENIX"));
            Teams.Add("CHELLOVEKK"                                  , new Team("CHELLOVEKK"));
            Teams.Add("FEARGGWP"                                    , new Team("FEARGGWP"));
            Teams.Add("&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;", new Team("&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"));
            Teams.Add("TAKA"                                        , new Team("TAKA"));

            Teams["BLACKSTAR98"].Winners        =   new ObservableCollection<Team>(){ Teams["FLEWLESS_PHOENIX"], Teams["FEARGGWP"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"], Teams["CHELLOVEKK"] };
            Teams["BLACKSTAR98"].Losers         =   new ObservableCollection<Team>(){ Teams["FLEWLESS_PHOENIX"], Teams["FEARGGWP"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"], Teams["CHELLOVEKK"] };

            Teams["FLEWLESS_PHOENIX"].Winners   =   new ObservableCollection<Team>(){ Teams["BLACKSTAR98"], Teams["FEARGGWP"], Teams["TAKA"] };
            Teams["FLEWLESS_PHOENIX"].Losers    =   new ObservableCollection<Team>(){ Teams["BLACKSTAR98"], Teams["FEARGGWP"], Teams["TAKA"], Teams["CHELLOVEKK"] };
            Teams["FLEWLESS_PHOENIX"].Tied      =   new ObservableCollection<Team>(){ Teams["CHELLOVEKK"] };

            Teams["CHELLOVEKK"].Winners         =   new ObservableCollection<Team>(){ Teams["FEARGGWP"], Teams["BLACKSTAR98"], Teams["FLEWLESS_PHOENIX"] };
            Teams["CHELLOVEKK"].Losers          =   new ObservableCollection<Team>(){ Teams["FEARGGWP"], Teams["BLACKSTAR98"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"] };
            Teams["CHELLOVEKK"].Tied            =   new ObservableCollection<Team>(){ Teams["TAKA"], Teams["FLEWLESS_PHOENIX"], Teams["&amp;#1058;&amp;#1040;&amp;#1050;&amp;#1040;"] };

            foreach (var t in Teams.Values)
                t.MakeLists();
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

            }
            catch (Exception ex) 
            {
                MessageBox.Show($"Ошибка при загрузке файла:\n{ex.Message}", MainControl?.Title);
                Log.Error(ex);
            }
        }
        private void            saveFileDialog      (object sender, ExecutedRoutedEventArgs e)
        {
            ////////////////////
            exportToExcel(ResultFileName);
            return;
            ////////////////////

            var dlg         =   new SaveFileDialog(){ Filter = "Excel files (*.xlsx)|*.xlsx" };

            if (dlg.ShowDialog() != true || string.IsNullOrEmpty(dlg.FileName))
                return;

            ResultFileName  =   dlg.FileName;
        }
        private void            exportToExcel       (string fileName)
        {
            if (File.Exists(fileName))
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении файла:\n{ex.Message}", MainControl?.Title);
                    return;
                }

            const   int     MIN_HEIGHT  =   4;
            const   int     MIN_WIDTH   =   19;
            const   int     START_ROW   =   3;
            const   int     START_COL   =   1;
            const   int     CHAR_DELTA  =   (byte)'A';
            const   char    START_COL_  =   (char)(START_COL + CHAR_DELTA);

            using (var  p       =   new ExcelPackage())
            {
                var     ws      =   p.Workbook.Worksheets.Add("Result");
                var     curRow  =   START_ROW;
                var     curCol  =   START_COL_;
                var     width   =   Math.Max(MIN_WIDTH, Teams.Max(t => Math.Max(t.Value.WinnersAndLosers.Length + t.Value.OnlyWinners.Length, 
                                                                                t.Value.WinnersAndLosers.Length + t.Value.OnlyLosers.Length)));

                Action<Team, int, int>
                    makeWinnersAndLosersCells   =   (t, r, c) =>
                {
                    for (var i = 0; i < t.WinnersAndLosers.Length; i++)
                    {
                        ws.Cells[$"{(char)(c + i)}{r}"].Value   =   t.WinnersAndLosers[i];
                        ws.Cells[$"{(char)(c + i)}{r}"].SetBackgroundColor(_WINERSLOSERS_COLOR);
                    }
                };
                Action<Team, int, int, bool>
                    makeWinnersOrLosersCells    =   (t, r, c, isWinners) =>
                {
                    var arr =   isWinners ? t.OnlyWinners : t.OnlyLosers;
                    for (var i = 0; i < arr.Length; i++)
                    {
                        ws.Cells[$"{(char)(width + START_COL - i + CHAR_DELTA - 1)}{r}"].Value   =   arr[i];
                        ws.Cells[$"{(char)(width + START_COL - i + CHAR_DELTA - 1)}{r}"].SetBackgroundColor(_WINERS_COLOR);
                    }
                };
                Action<Team, int, char, int>
                    makeTeamCells               =   (t, r, c, h) =>
                {
                    var     range                               =   $"{c}{r}:{(char)(c + width - 1)}{r + h - 1}";
                    ws.Cells[range].Merge                       =   true;
                    ws.Cells[range].Style.Font.Bold             =   true;
                    ws.Cells[range].Style.Font.Size             =   12;
                    ws.Cells[range].Style.HorizontalAlignment   =   OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    ws.Cells[range].Style.VerticalAlignment     =   OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    ws.Cells[$"{c}{r}"].Value                   =   t.Name;
                    ws.Cells[range].SetBackgroundColor(t.IsBettable ? _TEAM_COLOR2 : _TEAM_COLOR1);
                };
                Action<Team, int, char>
                    makeTiedCells               =   (t, r, c) =>
                {
                    for (var i = 0; i < t.OnlyTied.Length; i++)
                    {
                        ws.Cells[$"{c}{r + i}"].Value   =   t.OnlyTied[i];
                        ws.Cells[$"{c}{r + i}"].SetBackgroundColor(_WINERS_COLOR);
                    }
                };

                try
                {
                    foreach (var kvp in Teams)
                    {
                        var height  =   Math.Max(MIN_HEIGHT, kvp.Value.OnlyTied.Length);

                        makeWinnersAndLosersCells(kvp.Value, curRow, curCol);
                        makeWinnersOrLosersCells(kvp.Value, curRow, curCol, true);

                        curRow     +=   1;
                        makeTeamCells(kvp.Value, curRow, curCol, height);
                        makeTiedCells(kvp.Value, curRow, (char)((byte)START_COL_ + width));

                        curRow     +=   height;
                        makeWinnersAndLosersCells(kvp.Value, curRow, curCol);
                        makeWinnersOrLosersCells(kvp.Value, curRow, curCol, false);

                        curRow     +=   2;
                    }

                    var end =   (byte)START_COL_ + width - CHAR_DELTA + 1;
                    for (var i = (byte)START_COL_ - CHAR_DELTA + 1; i <= end; i++)
                        ws.Column(i).AutoFit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при генерации файла:\n{ex.Message}", MainControl?.Title);
                }

                try
                {
                    p.SaveAs(new FileInfo(fileName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении файла:\n{ex.Message}", MainControl?.Title);
                    return;
                }
            }

            MessageBox.Show("Файл успешно сохранен", MainControl?.Title);
        }

        #endregion
    }

    public  static  class   ExcelExtensions
    {
        public  static void     SetBackgroundColor  (this ExcelRange range, Color color)
        {
            range.Style.Fill.PatternType    =   OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }
    }
}
