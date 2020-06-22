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
            Log.Init();
            AppDomain.CurrentDomain.UnhandledException     +=   (s,a) =>    Log.Error(a.ExceptionObject as Exception);
            AppDomain.CurrentDomain.FirstChanceException   +=   (s,a) =>    Log.Error(a.Exception as Exception);

            base.OnStartup(e);
        }
    }
}
