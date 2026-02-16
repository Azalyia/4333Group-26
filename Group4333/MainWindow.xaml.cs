using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SadrievaAzalyia4333 sadrieva = new SadrievaAzalyia4333();
            sadrieva.Show();
        }
    }
}