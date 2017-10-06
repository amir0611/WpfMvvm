using System;
using System.Windows;

namespace ExcelToCsvWpfBased
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainWindowViewModel myMainWindowViewModel;
        public MainWindow()
        {
            InitializeComponent();
            myMainWindowViewModel = this.MyGridPanel.DataContext as MainWindowViewModel;

            if (myMainWindowViewModel != null)
            {
                var convertToCsvCommand = (MainWindowCommand)myMainWindowViewModel.ConvertToCsvCommand;

                convertToCsvCommand.Excecuting += OnUpdateFromViewModel;
                convertToCsvCommand.Excecuted += OnUpdateFromViewModel;
            }
        }

        private void OnUpdateFromViewModel(object sender, MessageNotificationEventArgs e)
        {
            NotifyUser(e.Message, e.MessageType);
        }

        private void NotifyUser(string message, MessageType messageType)
        {
            if (messageType == MessageType.Notification)
            {
                MessageBox.Show(message, "Notification", MessageBoxButton.OK);
            }

            if (messageType == MessageType.Confirmation)
            {
                var messageBoxResult = MessageBox.Show(message, "Confirmation", MessageBoxButton.YesNoCancel);
            }
        }
    }
}
