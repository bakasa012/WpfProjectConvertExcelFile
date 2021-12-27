using System.Windows;

namespace WpfProjectConvertExcelFile
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            BootStrapper bootStrapper = new BootStrapper();
            bootStrapper.Run();
        }
    }
}
