using Microsoft.Win32;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace OtzariaConveter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            RunConverter();
        }

        async void RunConverter()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                
                Filter = "Word Documents|*.doc;*.docm;*.docx;*.dotx;*.dotm;*.dot;*.odt;*.rtf",
                Multiselect = true,
                Title = "אנא בחר קבצים להמרה"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                progressBar.IsIndeterminate = true;
                await Task.Run(() => { new WordDocConverter().ConvertWordDocument(openFileDialog.FileNames); });
            }
            Close();
        }
    }
}

