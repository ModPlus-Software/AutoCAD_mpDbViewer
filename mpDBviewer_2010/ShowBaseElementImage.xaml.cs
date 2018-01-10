using System.Windows.Input;
using ModPlusAPI.Windows.Helpers;

namespace mpDbViewer
{
    public partial class ShowBaseElementImage
    {
        public ShowBaseElementImage()
        {
            InitializeComponent();
            this.OnWindowStartUp();
        }

        private void ShowBaseElementImage_OnKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
                Close();
        }
    }
}
