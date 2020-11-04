namespace mpDBviewer
{
    using Autodesk.AutoCAD.Runtime;
    using mpDbViewer;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    public class PluginStarter
    {
        public static MainWindow Window;

        [CommandMethod("ModPlus", "mpDBviewer", CommandFlags.Modal)]
        public void Start()
        {
#if !DEBUG
            ModPlusAPI.Statistic.SendCommandStarting(new ModPlusConnector());
#endif
            if (Window == null)
            {
                Window = new MainWindow();
                Window.Closed += (sender, args) =>
                {
                    Window = null;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                };
            }

            if (Window.IsLoaded)
                Window.Activate();
            else
                AcApp.ShowModelessWindow(AcApp.MainWindow.Handle, Window);
        }
    }
}
