using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace TextEditor
{
    public partial class App : Application
    {
        private const string MutexName = "TextEditor_Unique_Application_Mutex";
        private Mutex mutex;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        protected override void OnStartup(StartupEventArgs e)
        {
            // Убиваем все старые процессы с таким же именем
            KillOldProcesses();

            // Проверка на уже запущенный экземпляр
            mutex = new Mutex(true, MutexName, out bool createdNew);

            if (!createdNew)
            {
                MessageBox.Show("Приложение уже запущено!", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);

                // Найти и активировать существующее окно
                ActivateExistingWindow();

                Current.Shutdown();
                return;
            }

            base.OnStartup(e);

            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void KillOldProcesses()
        {
            try
            {
                string currentProcessName = Process.GetCurrentProcess().ProcessName;
                int currentId = Process.GetCurrentProcess().Id;

                foreach (var process in Process.GetProcessesByName(currentProcessName))
                {
                    if (process.Id != currentId)
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(3000);
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        private void ActivateExistingWindow()
        {
            try
            {
                string currentProcessName = Process.GetCurrentProcess().ProcessName;
                int currentId = Process.GetCurrentProcess().Id;

                foreach (var process in Process.GetProcessesByName(currentProcessName))
                {
                    if (process.Id != currentId && process.MainWindowHandle != IntPtr.Zero)
                    {
                        if (IsWindowVisible(process.MainWindowHandle))
                        {
                            ShowWindow(process.MainWindowHandle, 9);
                            SetForegroundWindow(process.MainWindowHandle);
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                Exception ex = e.ExceptionObject as Exception;
                MessageBox.Show($"Критическая ошибка: {ex?.Message}\n\nПриложение будет закрыто.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Environment.Exit(1);
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {
                MessageBox.Show($"Ошибка: {e.Exception.Message}\n\nПриложение продолжит работу.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            }
            catch { }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                if (mutex != null)
                {
                    mutex.ReleaseMutex();
                    mutex.Dispose();
                }
            }
            catch { }

            Environment.Exit(0);
            base.OnExit(e);
        }
    }
}