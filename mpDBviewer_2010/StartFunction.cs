#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using Autodesk.AutoCAD.Runtime;
using mpDbViewer;
using ModPlusAPI;

namespace mpDBviewer
{
    public class MainFunction
    {
        public static MpDbviewerWindow Window;
        [CommandMethod("ModPlus", "mpDBviewer", CommandFlags.Modal)]
        public void StartMpDBviewer()
        {
            Statistic.SendCommandStarting(new Interface());
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
