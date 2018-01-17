using System;
using System.Collections.Generic;
using System.Linq;
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
using Microsoft.Win32;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Data;
using System.Threading;

namespace FewerInvoices
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string FileName = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnAlege_Click(object sender, RoutedEventArgs e)
        {
            new Thread(() =>
            {
                App.Current.Dispatcher.Invoke(delegate
                {
                    btnAlege.IsEnabled = false;
                    btnStart.IsEnabled = false;
                });
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel *.xlsx |*.xlsx";
                if (ofd.ShowDialog() == true)
                {
                    MainWindow.FileName = ofd.FileName;

                    App.Current.Dispatcher.Invoke(delegate
                    {
                        txtFisier.Text = ofd.FileName;
                    });
                    
                    
                    ShowSheets(ofd.FileName);
                }
                App.Current.Dispatcher.Invoke(delegate
                {
                    btnAlege.IsEnabled = true;
                    btnStart.IsEnabled = true;
                });
            }).Start();
        }

        private void ShowSheets(string file)
        {
            if (!File.Exists(file))
            {
                MessageBox.Show("Fisierul ales nu exista!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            App.Current.Dispatcher.Invoke(delegate
            {
                lstSheets.Items.Clear();
            }); 

            try
            {
                XSSFWorkbook wkBook = new XSSFWorkbook(new System.IO.FileStream(file, System.IO.FileMode.Open));
                App.Current.Dispatcher.Invoke(delegate
                {
                    for (int i = 0; i < wkBook.NumberOfSheets; i++)
                    {
                        lstSheets.Items.Add(wkBook.GetSheetAt(i).SheetName);
                    }
                });
                
                wkBook.Close();
            }
            catch (IOException ioex)
            {
                MessageBox.Show("Nu se poate deschide fisierul:\r\n" + ioex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            catch (ICSharpCode.SharpZipLib.Zip.ZipException zex)
            {
                MessageBox.Show("Fisierul nu are formatul corect:\r\n" + zex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

        }

        private void lstSheets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MainWindow.FileName == null)
            {
                return;
            }

            string sel = lstSheets.SelectedValue.ToString();

            new Thread(() =>
            {
                ShowHeadRows(MainWindow.FileName, sel);
            }).Start();            
            
        }

        private String GetCellValue(ICell cell)
        {
            if (cell.CellType == CellType.Numeric)
            {
                return cell.NumericCellValue.ToString();
            }
            else if (cell.CellType == CellType.Blank)
            {
                return "";
            }
            else if (cell.CellType == CellType.Boolean)
            {
                return cell.BooleanCellValue.ToString();
            }
            else if (cell.CellType == CellType.Error)
            {
                return "ERROR";
            }
            else if (cell.CellType == CellType.Formula)
            {
                return "Formula";
            }
            else if (cell.CellType == CellType.Unknown)
            {
                return "UNKNOWN";
            }
            return cell.StringCellValue;
        }

        private string GetCellHeader(int number)
        {
            if (number <= 25)
            {
                return System.Convert.ToChar(65 + number).ToString();
            }

            int cnt = number / 26;
            int res = number - (cnt * 26);

            return System.Convert.ToChar(64 + cnt).ToString()+ System.Convert.ToChar(65 + res).ToString();
        }

        private void ShowHeadRows(String file, String sheetName)
        {
            App.Current.Dispatcher.Invoke(delegate {
                btnAlege.IsEnabled = false;
                btnStart.IsEnabled = false;
            });
            List<List<string>> lst = new List<List<string>>();

            if (!File.Exists(file))
            {
                MessageBox.Show("Fisierul ales nu exista!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            App.Current.Dispatcher.Invoke(delegate {
                dataGrid.ItemsSource = null;
            });
            
            try
            {
                XSSFWorkbook wkBook = new XSSFWorkbook(new System.IO.FileStream(file, System.IO.FileMode.Open));
                ISheet sheet = wkBook.GetSheet(sheetName);

                int cellPos = 0;
                int rowPos = 0;
                int maxCol = 0;

                while (rowPos < 10)
                {
                    List<string> lstCells = new List<string>();
                    cellPos = 0;     
                    int lastCell = (sheet.GetRow(rowPos)==null)?0:sheet.GetRow(rowPos).LastCellNum;
                    while (cellPos < lastCell)
                    {
                        
                        string val = sheet.GetRow(rowPos).GetCell(cellPos)==null?"": GetCellValue(sheet.GetRow(rowPos).GetCell(cellPos));                       
                        lstCells.Add((val == null) ? "" : val.ToString());
                        cellPos++;
                    }
                    maxCol = Math.Max(maxCol, cellPos);
                    lst.Add(lstCells);
                    rowPos++;
                }

                PopulateColumns(maxCol);

                wkBook.Close();
                App.Current.Dispatcher.Invoke(delegate {
                    dataGrid.ItemsSource = DataTableFromList(maxCol, lst).DefaultView;
                });
                

            }
            catch (IOException ioex)
            {
                MessageBox.Show("Nu se poate deschide fisierul:\r\n" + ioex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            catch (ICSharpCode.SharpZipLib.Zip.ZipException zex)
            {
                MessageBox.Show("Fisierul nu are formatul corect:\r\n" + zex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            App.Current.Dispatcher.Invoke(delegate {
                btnAlege.IsEnabled = true;
                btnStart.IsEnabled = true;
            });
        }

        private DataTable DataTableFromList(int maxCol, List<List<string>> lst)
        {
            DataTable dt = new DataTable();

            for (int i = 0; i < Math.Min(maxCol, 674); i++)
            {
                dt.Columns.Add(GetCellHeader(i));
            }

            foreach (List<string> row in lst)
            {

                DataRow dRow = dt.NewRow();

                for (int c = 0; c < row.Count; c++)
                {
                    dRow[c] = row[c];
                }

                dt.Rows.Add(dRow);

            }

            return dt;
        }

        private void PopulateColumns(int maxCol)
        {
            App.Current.Dispatcher.Invoke(delegate
            {
                cmbInvoice.Items.Clear();
                cmbItem.Items.Clear();

                for (int i = 0; i < Math.Min(maxCol, 674); i++)
                {
                    cmbInvoice.Items.Add(GetCellHeader(i));
                    cmbItem.Items.Add(GetCellHeader(i));
                }
            });
        }

        private void SaveInvoices(List<Invoice> lst)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Excel *.xlsx |*.xlsx";

            if (sfd.ShowDialog() != true)
            {
                return;
            }

            XSSFWorkbook wkBook = new XSSFWorkbook();
            ISheet shFacturi = wkBook.CreateSheet("Facturi");
            ISheet shItems = wkBook.CreateSheet("Items");

            int iRowFacturi = 0;
            int iRowItems = 0;

            foreach(Invoice inv in lst)
            {
                foreach(Item item in inv.Items)
                {
                    item.Visited = false;
                }
            }

            foreach (Invoice inv in lst)
            {
                IRow rFacturi = shFacturi.CreateRow(iRowFacturi++);
                rFacturi.CreateCell(0).SetCellValue(inv.Name);

                foreach (Item item in inv.Items)
                {
                    if (!item.Visited)
                    {
                        IRow rItems = shItems.CreateRow(iRowItems++);
                        rItems.CreateCell(0).SetCellValue(inv.Name);
                        rItems.CreateCell(1).SetCellValue(item.Name);
                        item.Visited = true;
                    }

                }
            }

            try
            {
                FileStream fs = File.Create(sfd.FileName);

                wkBook.Write(fs);

                fs.Close();

                MessageBox.Show("Gata","Info",MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Nu am putut salva fisierul excel: \r\n"+ex.Message, "Atentie!", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            if(lstSheets.SelectedIndex < 0)
            {
                MessageBox.Show("Alegeti va rog un sheet!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            if (cmbInvoice.SelectedIndex < 0)
            {
                MessageBox.Show("Alegeti va rog o coloana pentru facturi!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            if (cmbItem.SelectedIndex < 0)
            {
                MessageBox.Show("Alegeti va rog o coloana pentru items!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            btnStart.IsEnabled = false;
            btnAlege.IsEnabled = false;

            string file = txtFisier.Text;
            string sheetName = lstSheets.SelectedValue.ToString();
            int invoiceIndex = cmbInvoice.SelectedIndex;
            int itemIndex = cmbItem.SelectedIndex;
            statusBarIt.Content = "Import date din excel";

            new Thread(() => {

                if (!File.Exists(file))
                {
                    MessageBox.Show("Fisierul ales nu exista!", "Atentie", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                XSSFWorkbook wkBook = null;
                try
                {                    

                    wkBook = new XSSFWorkbook(new System.IO.FileStream(file, System.IO.FileMode.Open));
                    ISheet sheet = wkBook.GetSheet(sheetName);

                    int firstRow = firstRow = sheet.FirstRowNum;
                    int maxPgBar = sheet.LastRowNum - sheet.FirstRowNum;
                    
                    App.Current.Dispatcher.Invoke(
                        delegate
                        {
                            pgBar.Value = 0;
                            pgBar.Minimum = 0;
                            if (checkBox.IsChecked.HasValue && checkBox.IsChecked.Value)
                            {
                                firstRow++;
                                maxPgBar--;
                            }
                            pgBar.Maximum = maxPgBar;
                        }
                    );

                    int iRow = firstRow;
                    List<Item> lstItems = new List<Item>();
                    List <Invoice> lstInvoices = new List<Invoice>();

                    while (iRow <= sheet.LastRowNum)
                    {

                        Item item = new Item(GetCellValue(sheet.GetRow(iRow).Cells[itemIndex]));
                        Invoice invoice = new Invoice(GetCellValue(sheet.GetRow(iRow).Cells[invoiceIndex]));

                        Item foundItem = lstItems.Find((it) => it.Name.Equals(item.Name));
                        Invoice foundInvoice = lstInvoices.Find((inv) => inv.Name.Equals(invoice.Name));

                        if (foundItem != null)
                        {
                            item = foundItem;                        
                        }

                        if(foundInvoice != null)
                        {
                            invoice = foundInvoice;
                        }

                        item.AttemptAdd(invoice);
                        invoice.AttemptAdd(item);

                        if (foundItem == null)
                        {
                            lstItems.Add(item);
                        }                      
                        
                        if(foundInvoice == null)
                        {
                            lstInvoices.Add(invoice);
                        }  

                        App.Current.Dispatcher.Invoke(delegate
                        {
                            pgBar.Value = iRow++;                           
                        });

                    }

                    App.Current.Dispatcher.Invoke(delegate
                    {
                        statusBarIt.Content = "Procesez facturile";
                    });

                    List<Invoice> lstInvFinal =  Reducer.Reduce(lstInvoices, lstItems);

                    MessageBox.Show(String.Format("{0} Facturi initial, {1} Facturi final",lstInvoices.Count, lstInvFinal.Count));

                    SaveInvoices(lstInvFinal);                    
                    
                }
                catch (IOException ioex)
                {
                    MessageBox.Show("Nu se poate deschide fisierul:\r\n" + ioex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                catch (ICSharpCode.SharpZipLib.Zip.ZipException zex)
                {
                    MessageBox.Show("Fisierul nu are formatul corect:\r\n" + zex.Message, "Eroare", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                finally
                {
                    if (wkBook != null)
                    {
                        try { wkBook.Close(); } catch (Exception) { };
                    }
                    App.Current.Dispatcher.Invoke(delegate
                    {
                        pgBar.Value = 0;
                        statusBarIt.Content = "Mai putine facturi";
                    });
                }

                App.Current.Dispatcher.Invoke(delegate
                {
                    btnStart.IsEnabled = true;
                    btnAlege.IsEnabled = true;
                });

                }).Start();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Fewer Invoices. Progam pentru selectarea numarului minim de facturi \r\n care sa cuprinda toate itemurile cel putin o data.\r\n"+
                "\r\n"+
                "Create de Adrian Roland Clepcea. Pentru Hanes Brands.\r\n"+
                "Hanes Brands detine drepturile de autor!","Despre", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
