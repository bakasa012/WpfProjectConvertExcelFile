using System;
using System.Windows;
using Prism.Unity;

namespace WpfProjectConvertExcelFile
{
    public class BootStrapper: UnityBootstrapper
    {
        public override void Run(bool runWithDefaultConfiguration)
        {
            base.Run(runWithDefaultConfiguration);
        }

        protected override DependencyObject CreateShell()
        {
            return Container.TryResolve<Shell>();
        }

        protected override void InitializeShell()
        {
            App.Current.MainWindow = (Window)Shell;
            App.Current.MainWindow.Show();
        }

        protected override void ConfigureModuleCatalog()
        {
            base.ConfigureModuleCatalog();
            Type ModuleLocatorType = typeof(ModuleConvertExcelFile.ModuleLocators);
            ModuleCatalog.AddModule(new Prism.Modularity.ModuleInfo
            {
                ModuleName = ModuleLocatorType.Name,
                ModuleType = ModuleLocatorType.AssemblyQualifiedName
            });
        }
    }
}
