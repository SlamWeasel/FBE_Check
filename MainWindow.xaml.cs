using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace FBE_MonatsCheck
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class FBE_MonatMain : Window
    {
        string connectionString = @"secret";
        int lineCursor = 1;
        List<KeyValuePair<int, string>> resultAmount = new List<KeyValuePair<int, string>>();
        List<List<string>> resultSingle = new List<List<string>>();
        string blacklist = File.ReadAllText(Directory.GetCurrentDirectory() + @"\Blacklist.txt");

        public FBE_MonatMain()
        {
            InitializeComponent();
        }

        private void Execute_Click(object sender, RoutedEventArgs e)
        {
            if( Min.SelectedDate == Max.SelectedDate || 
                Min.SelectedDate > Max.SelectedDate ||
                Min.SelectedDate == null ||
                Max.SelectedDate == null)
                return;
            OutputTextBox.Text = "Abfragen werden ausgeführt (1/2)...";
            DateTime start = (DateTime)Min.SelectedDate,
                        end = ((DateTime)Max.SelectedDate).AddDays(1);

            resultAmount.Clear();
            resultSingle.Clear();


            SqlConnection cnn = new SqlConnection(connectionString);

            cnn.Open();

            SqlCommand amountCommand = new SqlCommand("" +
                @"  SELECT ...
                    -- für Zeitraum
                    st.Erstdat >= '" + start.ToString("dd.MM.yyyy") + @"' and st.Erstdat <= '" + end.ToString("dd.MM.yyyy") + @"'
                    
                    ... st.ErstUs not in " + blacklist +
                    "..." +

                    "GROUP BY st.ErstUs ORDER BY anzahl", cnn);
            amountCommand.CommandTimeout = 6000;
            SqlDataReader readAmount = amountCommand.ExecuteReader();
            while(readAmount.Read())
                resultAmount.Add(new KeyValuePair<int, string>(readAmount.GetInt32(0), readAmount.GetString(1)));
            cnn.Close();


            OutputTextBox.Text = "Abfragen werden ausgeführt (2/2)...";


            cnn = new SqlConnection(connectionString);

            cnn.Open();

            SqlCommand singleCommand = new SqlCommand("" +
                @"  SELECT  ...
                    --Monatlich
                    st.Erstdat >= '" + start.ToString("dd.MM.yyyy") + @"' and st.Erstdat <= '" + end.ToString("dd.MM.yyyy") + @"'
                     
                     ... and st.ErstUs not in " + blacklist +
                     " ..." +
                     "...", cnn);
            singleCommand.CommandTimeout =  6000;
            SqlDataReader readSingle = singleCommand.ExecuteReader();
            while (readSingle.Read())
                resultSingle.Add(new List<string> 
                { 
                    readSingle.GetString(0), 
                    readSingle.GetString(1), 
                    readSingle.GetDateTime(2).ToString("u"), 
                    Convert.ToString(readSingle.GetInt16(3)) 
                });
            cnn.Close();


            Excel.Application xlApp = new Excel.Application();

            xlApp.DisplayAlerts = false;

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheetAmount;
            Excel.Worksheet xlWorkSheetSingle;

            xlWorkBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            xlWorkSheetAmount = (Excel.Worksheet)xlWorkBook.Sheets.Item[1];
            xlWorkSheetAmount.Name = "Gesamtanzahl";
            xlWorkSheetSingle = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
            xlWorkSheetSingle.Name = "EinzelListe";

            OutputTextBox.Text = "Anzahl\tBenutzer\n" +
                                    "------------------------------------------------\n";
            xlWorkSheetAmount.Cells[1, 1] = "Anzahl";
            xlWorkSheetAmount.Cells[1, 2] = "Benutzer";
            lineCursor = 2;
            foreach (KeyValuePair<int, string> line in resultAmount)
            {
                xlWorkSheetAmount.Cells[lineCursor, 1] = line.Key;
                xlWorkSheetAmount.Cells[lineCursor, 2] = line.Value;
                OutputTextBox.Text += line.Key + "\t" + line.Value + '\n';
                lineCursor++;
            }
            OutputTextBox.Text += "\n\n\nBenutzer\t\tFERN-Nummer\t\tErstellt\t\tStatus\n" +
                                    "--------------------------------------------------------------------------\n";
            xlWorkSheetSingle.Cells[1, 1] = "Benutzer";
            xlWorkSheetSingle.Cells[1, 2] = "FERN-Nummer";
            xlWorkSheetSingle.Cells[1, 3] = "Erstellt";
            xlWorkSheetSingle.Cells[1, 4] = "Status-Nummer";
            lineCursor = 2;
            foreach (List<string> line in resultSingle)
            {
                int rowCursor = 1;
                foreach (string row in line)
                {
                    xlWorkSheetSingle.Cells[lineCursor, rowCursor] = row;
                    rowCursor++;
                }
                OutputTextBox.Text += line[0] + (line[0].Length <= 8 ? "\t\t" : "\t") + line[1] + '\t' + Convert.ToDateTime(line[2]).ToString("yyyy.MM.dd hh:mm") + '\t' + line[3] + '\n';
                lineCursor++;
            }


            try
            {
                string fileName = @"\\*Filepath*\Auswertungen\" + end.Year + @"\täglich\FBE\" + getMonthName(end.Month) + @"\FBE-" + start.ToString("dd.MM") + '-' + end.AddDays(-1).ToString("dd.MM") + ".xls";

                xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, null, null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                xlWorkBook.Close();
                xlApp.Quit();
                while (xlApp.Quitting) { }

                Marshal.ReleaseComObject(xlWorkSheetAmount);
                Marshal.ReleaseComObject(xlWorkSheetSingle);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                MessageBox.Show("Datei unter " + fileName + " abgespeichert", "Erfolgreich exportiert!");
            }
            catch (AccessViolationException)
            {
                MessageBox.Show("Sie haben keinen Zugriff auf diesen Ordner!\nDie Datei konnte nicht abgespeichert werden");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Es ist ein Fehler aufgetreten und die Datei konnte nicht abgespeichert werden!\n\n\n" + ex.ToString());
            }
        }

        public string getMonthName(int month)
        {
            switch (month)
            {
                case 01: return "01 Januar";
                case 02: return "02 Februar";
                case 03: return "03 März";
                case 04: return "04 April";
                case 05: return "05 Mai";
                case 06: return "06 Juni";
                case 07: return "07 Juli";
                case 08: return "08 August";
                case 09: return "09 September";
                case 10: return "10 Oktober";
                case 11: return "11 November";
                case 12: return "12 Dezember";
                default: return "00 Fehler";
            }
        }
    }
}
