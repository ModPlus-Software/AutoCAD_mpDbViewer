using System.Windows.Input;
using mpSettings;
using ModPlus;

namespace mpDbViewer
{
    /// <summary>
    /// Логика взаимодействия для ShowBaseElementImage.xaml
    /// </summary>
    public partial class ShowBaseElementImage
    {
        public ShowBaseElementImage()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
        }

        private void ShowBaseElementImage_OnKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
                this.Close();
        }
    }
}
