using System.Windows;

namespace Bobrus.App;

public partial class DefenderTipWindow : Window
{
    public DefenderTipWindow()
    {
        InitializeComponent();
    }

    private void OnCloseClicked(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
