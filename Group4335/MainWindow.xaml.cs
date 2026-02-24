using System.Windows;

namespace Group4335
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void AuthorInfoButton_Click(object sender, RoutedEventArgs e)
        {
            var infoWindow = new _4335_Nikulina(); 
            infoWindow.ShowDialog(); 
        }
    }
}