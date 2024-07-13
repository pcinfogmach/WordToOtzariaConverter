using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WordToOtzariaConverter
{
    static class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("aaa");
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents|*.doc;*.docx"
            };

            openFileDialog.ShowDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                ConvertWordDocument(openFileDialog.FileName);
            }
        }

        static void ConvertWordDocument(string filePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(filePath);
                string tempHtmlPath = Path.Combine(Path.GetTempPath(), "tempWordPreview.txt");
                wordDoc.SaveAs2(tempHtmlPath, WdSaveFormat.wdFormatFilteredHTML);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                // Close the Word document and quit the Word application
                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                }
                wordApp.Quit(false);
            }
        }
    }
}
