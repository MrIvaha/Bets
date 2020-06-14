using System;
using System.Windows;

namespace Ivaha.Bets
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected  override void    OnStartup   (StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException +=   (s,a) =>
            {
            ////MessageBox.Show($"Unhandled error:{Environment.NewLine}{a.ExceptionObject.ToString()}");
                Log.Error(a.ExceptionObject as Exception);
            };

            base.OnStartup(e);
        }
    }
}
