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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Threading;
using System.Collections.ObjectModel;
using System.Xml.Serialization;
using System.IO;

namespace Content
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        Excel.Range excelcells;
        Excel.Range excelcells1;
        Excel.Range excelcellsGroup;
        Excel.Range excelcellsParent;
        string pathfile;
        int posistion = 0;
        public static List<ExcelTypes> ListOfTypes;        
        List<TreeViewItem> treeViewItems;
        EditWindow editWindow;
        CatalogWindow catalogWindow;
        List<ExcelTypes> catalogsfromfile;
        /// <summary>
        /// Button for check erors and insert "Измерение_Характеристики" with empty values for geting data       
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Button1_Click(object sender, RoutedEventArgs e)
        {
            BusyIndicator.IsBusy = true;
            await Task.Run(() =>
            {
                //Create COM Objects. Create a COM object for everything that is referenced            
                xlWorksheet = xlWorkbook.Sheets[1];
            Start:
            xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;
            //insert "Измерение_Характеристики" previous "Значение_Характеристики"
            try
            {
                for (int i = 2; i <= colCount; i++)
                {
                    excelcells = (Excel.Range)xlWorksheet.Cells[1, i];
                    excelcells1 = (Excel.Range)xlWorksheet.Cells[1, i - 1];
                    if (Convert.ToString(excelcells.Value2) == "Значение_Характеристики" && Convert.ToString(excelcells1.Value2) != "Измерение_Характеристики")//Измерение_Характеристики
                    {
                        Read(excelcells.Address);//func                    
                        goto Start;// I like use goto ))
                    }
                }
            }
            catch { MessageBox.Show("Помилка!"); }
            //check for zero values in price where colunm name like Цена
            try
            {
                for (int i = 1; i <= colCount; i++)
                {
                    excelcells = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelcells.Value2) == "Цена")//Измерение_Характеристики
                    {
                        posistion = i;
                    }
                }
                if (posistion != 0)
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        excelcells = (Excel.Range)xlWorksheet.Cells[i, posistion];
                        if (Convert.ToString(excelcells.Value2) == "0")//Измерение_Характеристики
                        {
                            ((Range)xlWorksheet.Rows[i, Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
                            i = i - 2;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Помилка"); }
            });
            BusyIndicator.IsBusy = false;
        }
        /// <summary>
        /// function for replace empty name of field for Измерение_Характеристики
        /// </summary>
        /// <param name="address"></param>
        public void Read(string address)
        {
            Range rng;
            rng = xlWorksheet.get_Range(address, Missing.Value);
            rng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                                    XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            xlWorksheet.get_Range(address, Missing.Value).Value2 = "Измерение_Характеристики";
        }
        /// <summary>
        /// read and get data 
        /// equals catalog with real file data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private  void ReadFromFile_Click(object sender, RoutedEventArgs e)
        {           
              try
              {
                    xlWorksheet = xlWorkbook.Sheets[2];
                    //read from file real srtucture
                    xlRange = xlWorksheet.UsedRange;
                    int colCount = xlRange.Columns.Count;
                    int colRows = xlRange.Rows.Count;
                    ListOfTypes = new List<ExcelTypes>();

                    for (int j = 2; j <= colRows; j++)
                    {
                        excelcells1 = (Excel.Range)xlWorksheet.Cells[j, 2];
                        excelcellsGroup = (Excel.Range)xlWorksheet.Cells[j, 3];
                        excelcellsParent = (Excel.Range)xlWorksheet.Cells[j, 5];
                        if (Convert.ToString(excelcells1.Value2) != null)
                        {
                            ListOfTypes.Add(new ExcelTypes(Convert.ToString(excelcells1.Value2), Convert.ToString(excelcellsGroup.Value2), Convert.ToString(excelcellsParent.Value2)));
                        }
                    }                
            //  read from xml catalog
            XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));       
            try
            {
                using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.Open))
                {
                    List<ExcelTypes> newListWorkers = (List<ExcelTypes>)formatter.Deserialize(fs);
                    catalogsfromfile = newListWorkers;
                }
            }
            catch { MessageBox.Show("Файл CatalogContent.xml відсутній"); }
            //
            foreach (var listreal in ListOfTypes)
            {
                foreach (var listfile in catalogsfromfile)
                {
                    if (listreal.value == listfile.value)
                    {
                        listreal.Info = listfile.Info;
                    }
                }
            }
            //
            treeViewItems = new List<TreeViewItem>();
            foreach (var item in ListOfTypes)//create header for tree
            {
                if (item.ParentId == null)
                {
                    treeViewItems.Add(new TreeViewItem() { Header = item.value });
                }
            }
            //
            foreach (var item in ListOfTypes)//build tree
            {
                if (item.ParentId != null)
                {
                    foreach (var grid in ListOfTypes)
                    {
                        if (item.ParentId == grid.GroupId)
                        {
                            foreach (var ite in treeViewItems)
                                if (grid.value == Convert.ToString(ite.Header))
                                {
                                    ite.Items.Add(new TreeViewItem() { Header = item.value });
                                }
                        }
                    }
                }
            }
            foreach (var it in treeViewItems) { tree.Items.Add(it); }
              }
                catch { MessageBox.Show("Виберіть файл");}
        }

        public void GenerateCatalog()//dont using
        {
            XmlSerializer formatter = new XmlSerializer(typeof(List<Catalog>));
            catalogsfromfile = new List<ExcelTypes>();
            //deserialize
            try
            {
                using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.Open))
                {
                    List<ExcelTypes> newListWorkers = (List<ExcelTypes>)formatter.Deserialize(fs);
                    //View_Employers.ItemsSource = newListWorkers;
                    catalogsfromfile = newListWorkers;
                }
            }
            catch { MessageBox.Show("Файл CatalogContent.xml відсутній"); }
            //
            MessageBox.Show("'");
            /* if (File.Exists("CatalogContent.xml"))//temp delete for real view info
            {
                File.Delete("CatalogContent.xml");
            }

            using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, catalogs);
            }
            */
            //check for new catalogs names

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
            if (File.Exists("CatalogContent.xml"))//temp delete for real view info
            {
                File.Delete("CatalogContent.xml");
            }
            using (FileStream fs = new FileStream("CatalogContent.xml", FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, ListOfTypes);
            }
            string reservname = "ReservCatalogContent" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
            using (FileStream fs = new FileStream(reservname, FileMode.Create))
            {
                formatter.Serialize(fs, ListOfTypes);
            }
            MessageBox.Show("Saved");
        }
        /// <summary>
        /// add new element in list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem items = tree.SelectedItem as TreeViewItem;
            string selit = items.Header.ToString();
            ListBoxInfo.ItemsSource = null;
            foreach (var item in ListOfTypes)
            {
                if (selit == item.value)
                {
                    item.Info.Add("keyinfo");
                    ListBoxInfo.ItemsSource = item.Info;
                }
            }

        }
        /// <summary>
        /// add new item in listbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tree_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = tree.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in ListOfTypes)
                {
                    if (selit == item.value)
                    {
                        ListBoxInfo.ItemsSource = item.Info;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        /// <summary>
        /// delete selected item from listbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = tree.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in ListOfTypes)
                {
                    if (selit == item.value)
                    {
                        item.Info.RemoveAt(ListBoxInfo.SelectedIndex);
                        ListBoxInfo.ItemsSource = null;
                        ListBoxInfo.ItemsSource = item.Info;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент"); }
        }
        /// <summary>
        /// button for edit values from list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private  void ButtonEdit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = tree.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in ListOfTypes)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListBoxInfo.SelectedIndex].ToString() != null)
                        {
                               editWindow = new EditWindow(item.Info[ListBoxInfo.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                               editWindow.Show();
                               this.IsEnabled = false;                        
                               editWindow.Closed += ClosedEdit;                          
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        
        public void ClosedEdit (object eventt, System.EventArgs eventArgs)//delegate for event
            {
            try
            {
                TreeViewItem items = tree.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in ListOfTypes)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListBoxInfo.SelectedIndex].ToString() != null)
                        {
                            item.Info[ListBoxInfo.SelectedIndex] = EditWindow.infocontent;
                            ListBoxInfo.ItemsSource = null;
                            ListBoxInfo.ItemsSource = item.Info;
                        }
                    }
                }
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }
/// <summary>
/// Save
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
            if (File.Exists("Catalog\\CatalogContent.xml"))//temp delete for real view info
            {
                File.Delete("Catalog\\CatalogContent.xml");
            }
            using (FileStream fs = new FileStream("Catalog\\CatalogContent.xml", FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, ListOfTypes);
            }
            string reservname = "Catalog\\ReservCatalogContent" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
            using (FileStream fs = new FileStream(reservname, FileMode.Create))
            {
                formatter.Serialize(fs, ListOfTypes);
            }
            MessageBox.Show("Saved");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            MessageBox.Show("Done");
        }
        /// <summary>
        /// check for errors
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        
     
        private void MenuItem_Click_Exit(object sender, RoutedEventArgs e)
        {
            try { xlApp.Quit(); } catch { }
            System.Windows.Application.Current.Shutdown();
        }    


        private void MenuItem_Click_Help(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_Click_About(object sender, RoutedEventArgs e)
        {

        }
/// <summary>
/// Button fron menu
/// open and save path to file
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void MenuItem_Click_Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                pathfile = openFileDialog.FileName;
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(pathfile);
            }            
        }

        private void ListBoxInfo_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = tree.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in ListOfTypes)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListBoxInfo.SelectedIndex].ToString() != null)
                        {
                            editWindow = new EditWindow(item.Info[ListBoxInfo.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindow.Show();
                            this.IsEnabled = false;
                            editWindow.Closed += ClosedEdit;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        private delegate void UpdateProgressBarDelegate(System.Windows.DependencyProperty dp, Object value);       

        /// <summary>
        /// Put content from catalog in excell file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param> 
        private void ButtonPutContent_Click(object sender, RoutedEventArgs e)
        {            
          try
          {
                xlWorksheet = xlWorkbook.Sheets[1];               
                xlRange = xlWorksheet.UsedRange;
                int colCount = xlRange.Columns.Count;
                int rowCount = xlRange.Rows.Count;
                int position_Name = 0;//Название_позиции
                int position_Group = 0;//id of group
                int position_content = 0;//id content
                Excel.Range excelrange;
                Excel.Range excelrange_Name;
                Excel.Range excelrange_Group;
                Excel.Range excelrange_Content;
                Random rand = new Random();
                int val = 0;
            
            for (int i = 1; i <= colCount; i++)
                {
                    excelrange = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }
                    if (Convert.ToString(excelrange.Value2) == "Идентификатор_группы")
                    {
                        position_Group = i;
                    }
                    if (Convert.ToString(excelrange.Value2) == "Описание")
                    {
                        position_content = i;
                    }
                }
            ProgressBar1.Minimum = 0;
            ProgressBar1.Maximum = rowCount;
            ProgressBar1.Value = 0;
            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
            for (int i = 1; i <= rowCount; i++)
                  {
                    Dispatcher.Invoke(updatePbDelegate,System.Windows.Threading.DispatcherPriority.Background,new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    excelrange_Group = (Excel.Range)xlWorksheet.Cells[i, position_Group];
                    excelrange_Content = (Excel.Range)xlWorksheet.Cells[i, position_content];

                     foreach (var item in ListOfTypes)
                     {
                       if (Radio1.IsChecked == true)
                       { 
                          if (excelrange_Content.Value2 == null)
                          {
                             if (item.GroupId == Convert.ToString(excelrange_Group.Value2))
                             {
                                if (item.Info != null && item.Info.Count != 0)
                                {                             
                                    val = rand.Next(0, item.Info.Count);
                                    excelrange_Content.Value2 = Check_infokey(item.Info[val], Convert.ToString(excelrange_Name.Value2));
                                }
                             }
                          }
                       }
                    //
                    if (Radio2.IsChecked == true)
                    {                        
                            if (item.GroupId == Convert.ToString(excelrange_Group.Value2))
                            {
                                if (item.Info != null && item.Info.Count != 0)
                                {
                                    val = rand.Next(0, item.Info.Count);
                                    excelrange_Content.Value2 = Check_infokey(item.Info[val], Convert.ToString(excelrange_Name.Value2));
                                }
                            }                        
                    }
                }
                  }
            MessageBox.Show("Виконано"); 
          }
          catch { MessageBox.Show("Помилка, Можливо ви не відкрили файл"); }
        }
        public string Check_infokey(string key, string name)//replace keyinfo
        {
            string str;
            if (key.Contains("keyinfo"))
            { str = key.Replace("keyinfo", name); return str; }
            else { return key; }            
        }
        protected override void OnClosed(EventArgs e)
        {
            try { xlApp.Quit(); } catch { }            
            base.OnClosed(e);
        }

        private void MenuItem_Click_catalog_views(object sender, RoutedEventArgs e)
        {
            catalogWindow = new CatalogWindow();
            catalogWindow.Show();
           
        }
    }
}
