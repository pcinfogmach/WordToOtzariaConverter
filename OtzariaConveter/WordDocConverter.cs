using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OtzariaConveter
{
    public class WordDocConverter
    {
        public void ConvertWordDocument(string[] files)
        {
            foreach (var fileName in files)
            {
                Application wordApp = new Application();
                Document wordDoc = null;
                string extension = System.IO.Path.GetExtension(fileName);
                string outPutPath = fileName.Replace(extension, "_Otzaria.txt");
                try
                {
                    wordDoc = wordApp.Documents.Open(fileName);
                    wordDoc.SaveAs2(outPutPath, WdSaveFormat.wdFormatFilteredHTML);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("An error occurred: " + ex.Message);
                }

                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                }
                wordApp.Quit(false);

                ApplyReplacements(outPutPath);
            }
        }

        void ApplyReplacements(string filePath)
        {
            string content = File.ReadAllText(filePath);
            if (content.Contains("�"))
            {
                content = File.ReadAllText(filePath, Encoding.GetEncoding("Windows-1255"));
            }
            int indexOfBody = content.IndexOf("<body");
            if (indexOfBody > -1)
            {
                content = content.Substring(indexOfBody)
                    .Replace("</html>", "")
                    .Replace("class=MsoFootnoteText", "style='font-size:80%;'");
                content = Regex.Replace(content, "class=MsoList.*?(dir=RTL)[^>]+", "$1");
                content = Regex.Replace(content, "class=MsoList[^>]+", "$1");
                content = Regex.Replace(content, @"<.?body.*?>|<.?div.*?>", "");
                content = RemoveEmptyLines(content);
            }
            if (!string.IsNullOrEmpty(content)) { File.WriteAllText(filePath, content); }
        }

        static string RemoveEmptyLines(string input)
        {
            var lines = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var nonEmptyLines = lines.Where(line => !string.IsNullOrWhiteSpace(line));
            return string.Join("\n", nonEmptyLines);
        }
    }
}
