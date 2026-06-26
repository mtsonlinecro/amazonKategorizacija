using System.Diagnostics;

namespace AmazonCategoryMapperWinForms;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (_, e) => ShowAndLog(e.Exception, "WinForms thread exception");
            AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            {
                if (e.ExceptionObject is Exception ex) ShowAndLog(ex, "Unhandled application exception");
                else WriteLog("Unhandled application exception: " + Convert.ToString(e.ExceptionObject));
            };

            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            ShowAndLog(ex, "Startup exception");
        }
    }

    private static void ShowAndLog(Exception ex, string title)
    {
        var message = title + "\r\n\r\n" + ex;
        WriteLog(message);

        try
        {
            MessageBox.Show(
                "Aplikacija se srušila, ali sam spremio detalje u startup_error.log.\r\n\r\n" + ex.Message,
                "AmazonCategoryMapper - greška",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
        catch
        {
            // Ako MessageBox ne uspije, barem ostaje log datoteka.
        }
    }

    private static void WriteLog(string message)
    {
        try
        {
            var baseDir = AppContext.BaseDirectory;
            var logPath = Path.Combine(baseDir, "startup_error.log");
            File.AppendAllText(logPath,
                "==============================\r\n" +
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n" +
                message + "\r\n\r\n");

            // Kopija u projektni folder ako je moguće, da ga lakše nađeš.
            var dir = new DirectoryInfo(baseDir);
            while (dir != null)
            {
                var projectFile = Path.Combine(dir.FullName, "AmazonCategoryMapperWinForms.csproj");
                if (File.Exists(projectFile))
                {
                    File.AppendAllText(Path.Combine(dir.FullName, "startup_error.log"),
                        "==============================\r\n" +
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n" +
                        message + "\r\n\r\n");
                    break;
                }
                dir = dir.Parent;
            }
        }
        catch
        {
            // Zadnja linija obrane: ne ruši aplikaciju zbog logiranja.
        }
    }
}
