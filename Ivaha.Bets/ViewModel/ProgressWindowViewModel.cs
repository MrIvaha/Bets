using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;

namespace Ivaha.Bets.ViewModel
{
    public class ProgressWindowViewModel : ViewModelBase
    {
        public  event   EventHandler    OnCancel;

        public      ProgressWindowViewModel () : base (Application.Current?.MainWindow, CommandOwnerType)
        {
            RegCommand(Cancel, cancel);
        }

        #region Commands        |

        static  Type            CommandOwnerType    =   typeof(ProgressWindowViewModel);

        public  RoutedUICommand Cancel  { get; }    =   new RoutedUICommand("", "Cancel", CommandOwnerType);

        public  void            cancel  (object sender, ExecutedRoutedEventArgs e)
        {
            OnCancel?.Invoke(sender, EventArgs.Empty);
        }

        #endregion
    }
}
