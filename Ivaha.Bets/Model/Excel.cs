using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Ivaha.Bets.Model
{
    public  static  class   Excel
    {
        static readonly Color   _TEAM_COLOR1        =   Color.FromArgb(251, 228, 213);
        static readonly Color   _TEAM_COLOR2        =   Color.FromArgb(226, 239, 217);
        static readonly Color   _WINERSLOSERS_COLOR =   Color.FromArgb(248, 203, 172);
        static readonly Color   _WINERS_COLOR       =   Color.FromArgb(197, 224, 178);
        static readonly Color   _TIED_COLOR         =   Color.FromArgb(184, 204, 228);

        static readonly List<string>    _EXCLUSION_NAMES    =   new List<string>();
        static readonly Regex           _PATTERN_NAMES      =   new Regex(@".*\((.+)\).*");

                static          Excel               ()
        {
            _EXCLUSION_NAMES    =   Properties.Settings.Default.Exclusion.Split(new string[]{Environment.NewLine}, StringSplitOptions.None).Select(s => s.ToUpper()).ToList();
        }

        public  static  void    ImportFromExcel     (out IEnumerable<Team> teams, string fileName, string messageCaption, Action<string> callback)
        {
                teams   =   new List<Team>();
            var teams_  =   new Dictionary<string, Team>();

            Func<string, string>
                getName =   name =>
            {
                var match   =   _PATTERN_NAMES.Match(name);

                if (!match.Success || match.Groups.Count < 2)
                    return  name;

                return  _EXCLUSION_NAMES.Contains(match.Groups[1].Value.ToUpper()) ? name : match.Groups[1].Value;
            };

            Func<string, Team>
                addTeam =   name =>
            {
                if (teams_.ContainsKey(name))
                    return  teams_[name];

                var team=   new Team(name);
                teams_.Add(name, team);

                return  team;
            };

            Action<Team, Team, Func<Team, List<Team>>>
                addList =   (t1, t2, getList) =>
            {
                var lst =   getList(t1);
                if (!lst.Contains(t2))
                    lst.Add(t2);
            };

            if (!File.Exists(fileName))
                return;

            try
            {
                using (var p    =   new ExcelPackage(new FileInfo(fileName)))
                {
                     ExcelWorksheet  ws;

                    // Безумный код, но он работает
                    try     { ws=   p.Workbook.Worksheets.Count > 0 ? p.Workbook.Worksheets[1] : null; }
                    catch   { ws=   p.Workbook.Worksheets.Count > 0 ? p.Workbook.Worksheets[1] : null; }

                    callback?.Invoke("Excel документ успешно загружен");

                    // Todo
                    var startRow=   5;
                    var team1Col=   14;
                    var team2Col=   15;
                    var resCol  =   18;

                    if (ws.Dimension.End.Row < startRow || ws.Dimension.End.Column < resCol)
                    {
                        callback?.Invoke(" - Недостаточное количество строк или колонок в документе");
                        return;
                    }

                    callback?.Invoke(" - Требуемое количество строк и колонок найдено");

                    Team    team1;
                    Team    team2;

                    for(var row = startRow; row <= ws.Dimension.End.Row; row++)
                    {
                        team1   =   addTeam(getName(ws.Cells[row, team1Col].Text));
                        team2   =   addTeam(getName(ws.Cells[row, team2Col].Text));

                        if (!byte.TryParse(ws.Cells[row, resCol].Text, out var res))
                            continue;

                        switch (res)
                        {
                            case 0  :
                                addList(team1, team2, t => t.Tied);
                                addList(team2, team1, t => t.Tied);
                                break;
                            case 1  :
                                addList(team1, team2, t => t.Losers);
                                addList(team2, team1, t => t.Winners);
                                break;
                            case 2  :
                                addList(team1, team2, t => t.Winners);
                                addList(team2, team1, t => t.Losers);
                                break;
                            default : continue;
                        }
                    }
                }

                foreach (var team in teams_.Values)
                    team.MakeAllLists();

                teams           =   teams_.Select(kvp => kvp.Value).ToArray();

                callback?.Invoke($" - {teams_.Count} команд было обнаружено");
                callback?.Invoke($"   ...из которых на {teams.Count(t => t.IsBettable)} можно делать ставки{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                ShowMessage("Ошибка импорта из файла", messageCaption, ex);
            }

            ShowMessage("Файл успешно импортирован", messageCaption);
        }
        public  static  bool    ExportToExcel       (IEnumerable<Team> teams, string fileName, string messageCaption, Action<string> callback, CancellationToken token)
        {
            if (File.Exists(fileName))
                try
                {
                    File.Delete(fileName);
                }
                catch (Exception ex)
                {
                    ShowMessage("Ошибка при удалении существующего файла", messageCaption, ex);
                    return  false;
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
                var     width   =   Math.Max(MIN_WIDTH, teams.Max(t => Math.Max(t.WinnersAndLosers.Length + t.OnlyWinners.Length, 
                                                                                t.WinnersAndLosers.Length + t.OnlyLosers.Length)));

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
                        ws.Cells[$"{(char)(width + START_COL - i + CHAR_DELTA - 1)}{r}"].SetBackgroundColor(t.OnlyTied.Contains(arr[i]) ? _TIED_COLOR : _WINERS_COLOR);
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
                    ws.Cells[$"{c}{r}"].Value                   =   $"{t.Name})";
                    ws.Cells[range].SetBackgroundColor(t.IsBettable ? _TEAM_COLOR2 : _TEAM_COLOR1);
                };
                Action<Team, int, char>
                    makeTiedCells               =   (t, r, c) =>
                {
                    for (var i = 0; i < t.OnlyTied.Length; i++)
                    {
                        ws.Cells[$"{c}{r + i}"].Value   =   t.OnlyTied[i];
                        ws.Cells[$"{c}{r + i}"].SetBackgroundColor(_TIED_COLOR);
                    }
                };

                try
                {
                    foreach (var team in teams)
                    {
                        token.ThrowIfCancellationRequested();

                        var height  =   Math.Max(MIN_HEIGHT, team.OnlyTied.Length);

                        makeWinnersAndLosersCells(team, curRow, curCol);
                        makeWinnersOrLosersCells(team, curRow, curCol, true);

                        curRow     +=   1;
                        makeTeamCells(team, curRow, curCol, height);
                        makeTiedCells(team, curRow, (char)((byte)START_COL_ + width));

                        curRow     +=   height;
                        makeWinnersAndLosersCells(team, curRow, curCol);
                        makeWinnersOrLosersCells(team, curRow, curCol, false);

                        curRow     +=   2;
                    }

                    var end =   (byte)START_COL_ + width - CHAR_DELTA + 1;
                    for (var i = (byte)START_COL_ - CHAR_DELTA + 1; i <= end; i++)
                        ws.Column(i).AutoFit();
                }
                catch (OperationCanceledException)
                {
                    return  false;
                }
                catch (Exception ex)
                {
                    ShowMessage("Ошибка при генерации файла", messageCaption, ex);
                }

                try
                {
                    p.SaveAs(new FileInfo(fileName));

                    callback?.Invoke($"{ws.Dimension.End.Row} строк записано в файл{Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    ShowMessage("Ошибка при сохранении файла", messageCaption, ex);
                    return  false;
                }
            }

            return  true;
        }
        private static  void    ShowMessage         (string message, string messageCaption, Exception ex = null)    =>  
            MessageBox.Show($"{message}{(ex == null ? "" : $":{Environment.NewLine}{ex.Message}")}", messageCaption);
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
