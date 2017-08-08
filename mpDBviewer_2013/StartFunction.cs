using System;
using Autodesk.AutoCAD.Runtime;
using mpDbViewer;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace mpDBviewer
{
    public class MainFunction
    {
        public static MpDbviewerWindow Window;
        [CommandMethod("ModPlus", "mpDBviewer", CommandFlags.Modal)]
        public void StartMpDBviewer()
        {
            if (Window == null)
            {
                Window = new MpDbviewerWindow();
                Window.Closed += win_Closed;
            }

            if (Window.IsLoaded)
                Window.Activate();
            else
                AcApp.ShowModelessWindow(AcApp.MainWindow.Handle, Window);
        }

        static void win_Closed(object sender, EventArgs e)
        {
            Window = null;
            // Перевод фокуса
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
        }
    }
}
