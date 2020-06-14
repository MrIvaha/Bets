using Ivaha.Bets.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Ivaha.Bets
{
    public partial class ProgressWindow : Window
    {
        /// <summary>Запуск модального окна с индикацией процесса и выполнением процедуры в отдельном потоке
        /// </summary>
        /// <param name="taskAction">Процедура для отдельного потока</param>
        /// <param name="cts">Источник токена для взвода отмены</param>
        /// <param name="token">Токен для останова параллельного потока в случае отмены</param>
        /// <param name="cancelVisible">Показывать ли кнопку отмены</param>
        /// <param name="mode">Режим диалогового окна - с прогресс баром или бесконечной крутилкой</param>
        public  static  void            Run             (Action                                         // Процедура параллельной задачи
                                                         <
                                                            IProgress<(byte Percent, string Label)>,    // Для отрисовки прогресса из процедры
                                                            CancellationTokenSource,                    // Источник маркера отмены - по нему происходит отмена из самой процедры в случае подветрждения пользователем прервать скачивание в виду ошибки на одном из файлов
                                                            CancellationToken,                          // Сам маркер отмены - просто пробрасывается в процедуру, где по нему будет обрабатываться отмена
                                                            Window                                      // Окно будет являться родительским для всех модальных сообщений из процедуры
                                                         > taskAction, 
                                                         CancellationTokenSource cts,                   // Источник маркера отмены - по нему будет происходить отмена при нажатии в окне прогресса Отмена
                                                         CancellationToken token,                       // Маркер отмены для проброса в процедуру
                                                         Action<Exception, Window> exceptionCallback    // Колбэк для обработки ошибки (например для вывода сообщения об ошибке) из процедуры
                                                         )
        {
            var progress    =   new Progress<(byte Percent, string Label)>();   // Индикация процесса
            var window      =   new ProgressWindow();                           // Модальное окно ожидания

            // Инициализация окна (вью модели)
            if (window.DataContext is ProgressWindowViewModel model)
            {
                model.OnCancel     +=   (s,e) => 
                {
                    cts.Cancel();
                };
            }

            // При загрузке окна произойдет запуск параллельного потока, после выполнения которого окно закроется
            window.Loaded  +=   async (s,e) => 
            {
                try
                {
                    await Task.Run(() => taskAction(progress, cts, token, window), token);
                }
                catch (OperationCanceledException){ }
                catch (Exception ex){ exceptionCallback?.Invoke(ex, window); }
                finally
                {
                    if (s is ProgressWindow sender)
                    {
                        sender.TaskExecution    =   false;
                        sender.Close();
                    }
                }
            };

            window.ShowDialog();
        }

        public              bool        TaskExecution   { get; set; }   =   true;

        public                          ProgressWindow  ()
        {
            InitializeComponent();

            Closing    +=   (s,e) => e.Cancel = TaskExecution;   // Блокировка закрытия окна по кресту
            Closed     +=   (s,e) => 
            {
                if (DataContext is IDisposable disposable)
                    disposable.Dispose();
            };
        }
        // Todo Почему-то с привязанной командой кнопка залочена + почему-то выполнение Execute команды Cancel модели не приводит к вызову cancel...
        private             void        Button_Click    (object sender, RoutedEventArgs e)
        {
            if (DataContext is ProgressWindowViewModel model)
                model.cancel(null, null);
        }
    }
}
