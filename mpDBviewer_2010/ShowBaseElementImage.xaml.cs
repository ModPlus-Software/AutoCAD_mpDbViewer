using System.Windows.Input;

namespace mpDbViewer
{
    public partial class ShowBaseElementImage
    {
        public ShowBaseElementImage()
        {
            InitializeComponent();
        }

        private void ShowBaseElementImage_OnKeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape)
                Close();
        }
    }
}
