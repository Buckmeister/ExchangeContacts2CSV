using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using BCK;
using Microsoft.Win32;
using System.Windows.Threading;

namespace ExchangeContacts2CSV
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly DispatcherTimer tmAlignEditorWindow;
        private readonly EditorWindow editorWindow;
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();

            editorWindow = new EditorWindow();

            tmAlignEditorWindow = new DispatcherTimer
            {
                Interval = new TimeSpan(0, 0, 0, 0, 500),
                IsEnabled = false,
            };

            tmAlignEditorWindow.Tick += tmAlignEditorWindow_Tick;

            txtLog.Text =

@"
Bitte geben Sie Ihre, Zugangsdaten ein und melden sich an.
Wählen darauf einen Kontaktordner aus und betätigen Sie 'Importieren'.

Sie erhalten nun eine Vorschau auf die gefunden Kontakte, 
die Sie in dieser Anwendung vor dem Export editieren können.

Zum Schluss wählen Sie den Pfad zur Ausgabedatei aus und betätigen 
Sie die Schaltfläche 'Exportieren', um die Ausgabedatei zu befüllen.

";

            txtLog.Text += $"(Programmversion: {DeploymentAction.GetRunningVersion()})";
            
            Console.SetOut(new TextBoxOutputter(txtLog));
        }

        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Datei (*.csv)|*.csv",
                FileName = "Kontakte.csv"
            };

            if (saveFileDialog.ShowDialog() == true)
                txtPathToCsv.Text = saveFileDialog.FileName;
        }

        private void pbPassword_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (DataContext != null)
            { ((dynamic)DataContext).SecurePassword = ((PasswordBox)sender).SecurePassword; }
        }

        private void btnLogin_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount >= 2)
            {
                e.Handled = true;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            editorWindow.DataContext = DataContext;
            editorWindow.Owner = this;
            editorWindow.Align();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        void tmAlignEditorWindow_Tick(object sender, EventArgs e)
        {
            tmAlignEditorWindow.IsEnabled = false;
            editorWindow.Align();
        }

        private void InitiateEditorWindowAlignment()
        {
            tmAlignEditorWindow.IsEnabled = true;
            tmAlignEditorWindow.Stop();
            tmAlignEditorWindow.Start();
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            InitiateEditorWindowAlignment();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            InitiateEditorWindowAlignment();
        }
    }
}
