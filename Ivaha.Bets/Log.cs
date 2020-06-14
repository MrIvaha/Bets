using System;
using System.IO;
using System.Runtime.CompilerServices;
using log4net;
using log4net.Config;

namespace Ivaha.Bets
{
    public  static  class Log
    {
                static              Log         ()
        {
            var fileInfo    =   new FileInfo("Log4Net.config");

            if (fileInfo.Exists)
                XmlConfigurator.ConfigureAndWatch(fileInfo);
        }

        private static  ILog        logger          =   LogManager.GetLogger("Ivaha.Bets");

        public  static  void        Debug       (string message, Exception ex = null)   => logger.Debug (message, ex);
        public  static  void        Error       (Exception ex = null, 
                                                 [CallerMemberName] string memberName = "", 
                                                 [CallerFilePath] string filePath = "")     => logger?.Error ($"Exception in {memberName} in {System.IO.Path.GetFileName(filePath)}.", ex);
        public  static  void        Fatal       (string message, Exception ex = null)   => logger.Fatal (message, ex);
        public  static  void        Info        (string message, Exception ex = null)   => logger.Info  (message, ex);
        public  static  void        Warn        (string message, Exception ex = null)   => logger.Warn  (message, ex);
    }
}