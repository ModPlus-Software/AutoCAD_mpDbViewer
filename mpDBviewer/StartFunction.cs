namespace mpDBviewer
{
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;
    using mpDbViewer;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    public class MainFunction
    {
        public static MpDbviewerWindow Window;
        
        [CommandMethod("ModPlus", "mpDBviewer", CommandFlags.Modal)]
        public void StartMpDBviewer()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());
            if (Window == null)
            {
                Window = new MpDbviewerWindow();
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
