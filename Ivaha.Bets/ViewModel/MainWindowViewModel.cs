using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Win32;

namespace Ivaha.Bets.ViewModel
{
    public  class   MainWindowViewModel : ViewModelBase
    {
        public                  MainWindowViewModel () : base (Application.Current?.MainWindow, CommandOwnerType)
        {
            
        }

        #region Commands    |

        static  Type            CommandOwnerType    =   typeof(MainWindowViewModel);

        public RoutedUICommand  OpenFileDialog      { get; }    =   new RoutedUICommand("", "OpenFileDialog"    , CommandOwnerType);
        public RoutedUICommand  SaveFileDialog      { get; }    =   new RoutedUICommand("", "SaveFileDialog"    , CommandOwnerType);

        private void            openFileDialog      (object sender, ExecutedRoutedEventArgs e)
        {
        ////var dlg =   new OpenFileDialog(){ Filter = "Csv файлы Altium (*.csv)|*.csv|Kyinside файлы (*.kyinside)|*.kyinside" };
            
        ////if (dlg.ShowDialog() != true || string.IsNullOrEmpty(dlg.FileName) || !File.Exists(dlg.FileName))
        ////    return;

        ////OpenFileName                =   dlg.FileName;

        ////// Try load xls
        ////try
        ////{
                
        ////}
        ////catch (Exception ex) 
        ////{
        ////    MessageBox.Show($"Ошибка при загрузке файла:\n{ex.Message}", MainControl?.Title);
        ////    Log.Error(ex, new StackTrace(true).GetFrame(0).GetMethod().Name, new StackTrace(true).GetFrame(0).GetFileName());
        ////}
        }
        private void            saveFileDialog      (object sender, ExecutedRoutedEventArgs e)
        {
        ////var dlg =   new SaveFileDialog(){ Filter = "X файлы (*.x)|*.x" };
            
        ////if (dlg.ShowDialog() != true || string.IsNullOrEmpty(dlg.FileName))
        ////    return;

        ////SaveFileName    =   dlg.FileName;
        }

        #endregion
    }
}
