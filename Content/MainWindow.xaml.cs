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
using MahApps;
using System.Text.RegularExpressions;
using System.Windows.Threading;

namespace Content
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public partial class MainWindow : MahApps.Metro.Controls.MetroWindow
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
        public static string PathCatalog = "Catalog\\CatalogContent.xml";//default
        int posistion = 0;
        public static List<ExcelTypes> ListOfTypes;        
        List<TreeViewItem> treeViewItems;
        EditWindow editWindow;
        CatalogWindow catalogWindow;
        RestructWindow restructWindow;
        KeyWordsWindow KeyWordsWindow;
        ReStruct reStruct;
        CatalogWords catalogWords;
        RegexWindow regexWindow;
        List<ExcelTypes> catalogsfromfile;
        List<MainWords> MainWords;
        /// <summary>
        /// Button for check erors and insert "Измерение_Характеристики" with empty values for geting data       
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Button1_Click(object sender, RoutedEventArgs e)
        {
            try
            {               
                await Task.Run(() =>
                {
                try
                {
                    //Create COM Objects. Create a COM object for everything that is referenced            
                    xlWorksheet = xlWorkbook.Sheets[1];
                    Start:
                    xlRange = xlWorksheet.UsedRange;
                    int colCount = xlRange.Columns.Count;
                    int rowCount = xlRange.Rows.Count;

                    //insert "Измерение_Характеристики" previous "Значение_Характеристики"

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
                    //check zero values
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
                    catch { MessageBox.Show("Виберіть файл для роботи!"); }
                });           
               
            }
            catch { MessageBox.Show("Виберіть файл для роботи!"); }
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
            try
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
            catch { MessageBox.Show("Виберіть файл для роботи!");}
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
        public void ClosedCatalog(object eventt, System.EventArgs eventArgs)//delegate for event Catalog
        {
            catalogsfromfile = null;
            tree.Items.Clear();
            ListBoxInfo.ItemsSource = null;
            //ListBoxInfo.Items.Clear();
            XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
            try
            {
                using (FileStream fs = new FileStream(PathCatalog, FileMode.Open))
                {
                    List<ExcelTypes> newListWorkers = (List<ExcelTypes>)formatter.Deserialize(fs);
                    catalogsfromfile = newListWorkers;
                }
            }
            catch { MessageBox.Show("Файл " + PathCatalog + "відсутній"); }
            try
            {
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
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Ви не відкрили файл каталог!"); this.IsEnabled = true; }
        }
        /// <summary>
        /// Save
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
                string reservname = "Catalog\\ReservCatalogContent" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
                if (ListOfTypes != null)
                {
                    using (FileStream fs = new FileStream(reservname, FileMode.Create))
                    {
                        formatter.Serialize(fs, ListOfTypes);
                    }
                    if (File.Exists("Catalog\\CatalogContent.xml"))//temp delete for real view info
                    {
                        File.Delete("Catalog\\CatalogContent.xml");
                    }
                    using (FileStream fs = new FileStream("Catalog\\CatalogContent.xml", FileMode.OpenOrCreate))
                    {
                        formatter.Serialize(fs, ListOfTypes);
                    }
                    MessageBox.Show("Резервна копія каталога збережена\n " + reservname + "");                    
                }
                else { MessageBox.Show("Ви не відкрили каталог для роботи!"); }
            }
            catch { MessageBox.Show("Помилка"); }
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
        /// update data and view of tree
        /// search similar data in catalog and real file
        /// </summary>
        /// <param name="excelTypes"></param>
        public void UpdateData()
        {
            catalogsfromfile = null;
            tree.Items.Clear();
            ListBoxInfo.Items.Clear();
            XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
            try
            {
                using (FileStream fs = new FileStream(PathCatalog, FileMode.Open))
                {
                    List<ExcelTypes> newListWorkers = (List<ExcelTypes>)formatter.Deserialize(fs);
                    catalogsfromfile = newListWorkers;
                }                
            }
            catch { MessageBox.Show("Файл " + PathCatalog + "відсутній"); }
            foreach (var listreal in ListOfTypes)
            {
                foreach (var listfile in catalogsfromfile)
                {
                    if (listreal.value == listfile.value)
                    {
                        listreal.Info = listfile.Info;
                        listreal.keyWords = listfile.keyWords;
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
                        //grid.Keys = new List<List<string>>();
                        //grid.Keys.Add(new List<string> { "1","2","3"});
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
       
/// <summary>
/// Button fron menu
/// open and save path to file
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        public void MenuItem_Click_Open(object sender, RoutedEventArgs e)
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
                UpdateData();                 
            }
            catch { MessageBox.Show("Виберіть файл"); }
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
        //
        public string Check_infoword(string key, string name, string article)//replace keyinfo
        {
            string[] article1 = { article };
            if (article == null) { return key.Replace("keyword", name); }
            else
            if (key.Contains("keyword") && name.Contains(article))
            {
                string[] result = name.Split(article1, StringSplitOptions.RemoveEmptyEntries);
                key = key.Replace("keyword", result[0]);
                return key;
            }
            else { return key.Replace("keyword", name); }
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
            this.IsEnabled = false;
            catalogWindow.Closed += ClosedCatalog;
        }

        private void SaveInFile_Click(object sender, RoutedEventArgs e)
        {
           // try
           // {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);               
         //   }
          //  catch { MessageBox.Show("Файл не відкритий!"); }
        }

        private void Button_viev_words_Click(object sender, RoutedEventArgs e)
        {
            catalogWords = new CatalogWords();
            catalogWords.Show();
        }

        private void GenerateWords_Click(object sender, RoutedEventArgs e)
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
                Range excelrange;
                Excel.Range excelrange_Name;
                Excel.Range excelrange_Group;
                Excel.Range excelrange_Content;
                Excel.Range excelrange_ContentWords;
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
                    //if (Convert.ToString(excelrange.Value2) == "Название_Характеристики")
                    //{
                    //    position_content = i;
                    //}
                }
                ProgressBar1.Minimum = 0;
                ProgressBar1.Maximum = rowCount;
                ProgressBar1.Value = 0;
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                val = position_Group+1;//=Название_Характеристики
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    excelrange_Group = (Excel.Range)xlWorksheet.Cells[i, position_Group];
                    position_content = val;
                    pos1:
                    excelrange_Content = (Excel.Range)xlWorksheet.Cells[i, position_content];
                    excelrange_ContentWords = (Excel.Range)xlWorksheet.Cells[i, position_content+2];
                    string Name = Convert.ToString(excelrange_Name.Value2);
                    foreach (var item in ListOfTypes)
                    {                                                
                            if (excelrange_Content.Value2 == null)
                            {
                                if (item.GroupId == Convert.ToString(excelrange_Group.Value2))
                                {
                                    if (item.keyWords != null)
                                    {
                                        foreach (var el in item.keyWords)
                                        {
                                            if (el.ListValues != null)
                                            {
                                                foreach (var words in el.ListValues)
                                                {
                                                    if (Name.Contains(words.Replace("\n",""))==true)
                                                    {
                                                        
                                                    Name= Name.Replace(words.Replace("\n", ""), "");
                                                    excelrange_Name.Value2 = "";
                                                    excelrange_Name.Value2 = Name;
                                                        excelrange_Content.Value2 = el.KeyName.Replace("\n", "");
                                                        excelrange_ContentWords.Value2 = words.Replace("\n", "");
                                                        position_content = position_content + 3;
                                                        excelrange_Content = (Excel.Range)xlWorksheet.Cells[i, position_content];
                                                        excelrange_ContentWords = (Excel.Range)xlWorksheet.Cells[i, position_content + 2];
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                            }
                            else
                            {
                                position_content = position_content + 3;
                                goto pos1;
                            }   
                    }
                }
                MessageBox.Show("Виконано");
            }
            catch { MessageBox.Show("Помилка, Можливо ви не відкрили файл"); }

        }
        /// <summary>
        /// Automatic generate 
        /// 1- check for errors 
        /// 2- parse colums insert null colums previous Значение_Характеристики
        /// 3- find words from catalog and create characteristics
        /// 4- write marketing using catalog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AutoGenerate_Click(object sender, RoutedEventArgs e)
        {
            DateTime timestart = DateTime.Now;
            ProgressBar1.Minimum = 0;            
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;
            ProgressBar1.Maximum = colCount;
            try
            {
                //1
                //Create COM Objects. Create a COM object for everything that is referenced      
                ProgressBar1.Value = 0;
                Start:
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                //insert "Измерение_Характеристики" previous "Значение_Характеристики"

                for (int i = 2; i <= colCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelcells = (Excel.Range)xlWorksheet.Cells[1, i];
                    excelcells1 = (Excel.Range)xlWorksheet.Cells[1, i - 1];
                    if (Convert.ToString(excelcells.Value2) == "Значение_Характеристики" && Convert.ToString(excelcells1.Value2) != "Измерение_Характеристики")//Измерение_Характеристики
                    {
                        Read(excelcells.Address);//func                    
                        goto Start;// I like use goto ))
                    }
                }
                //2
                label1.Background=Brushes.Green;
                ProgressBar1.Value = 0;
                ProgressBar1.Maximum = rowCount;
                for (int i = 1; i <= colCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
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
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                        excelcells = (Excel.Range)xlWorksheet.Cells[i, posistion];
                        if (Convert.ToString(excelcells.Value2) == "0")//Измерение_Характеристики
                        {
                            ((Range)xlWorksheet.Rows[i, Missing.Value]).Delete(XlDeleteShiftDirection.xlShiftUp);
                            i = i - 2;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Виберіть файл для роботи!"); }
            label2.Background = Brushes.Green;
            ProgressBar1.Value = 0;
            ProgressBar1.Maximum = rowCount;
            //regex            
                List<RegexC> catalogRegex;
                XmlSerializer formatter = new XmlSerializer(typeof(List<RegexC>));
                using (FileStream fs = new FileStream("Catalog\\catalogRegexc.xml", FileMode.Open))
                {
                    catalogRegex = (List<RegexC>)formatter.Deserialize(fs);
                }
                label5_Copy.Content = catalogRegex.Count-1;
                #region set regex values 
                int position_Name9 = 0;//Название_позиции
                int position_Group9 = 0;//id of group
                int position_content9 = 0;//id content
                Excel.Range excelrange29;
                Excel.Range excelrange_Name29;
                Excel.Range excelrange_Group29;
                Excel.Range excelrange_Content29;
                Excel.Range excelrange_ContentWords29;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange29 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange29.Value2) == "Название_позиции")
                    {
                        position_Name9 = i;
                    }
                    if (Convert.ToString(excelrange29.Value2) == "Идентификатор_группы")
                    {
                        position_Group9 = i;
                    }
                    if (Convert.ToString(excelrange29.Value2) == "Название_Характеристики")
                    {
                        position_content9 = i;
                        goto f1;
                    }
                }
                f1:
                ProgressBar1.Maximum = rowCount;
                for (int j = 0; j < catalogRegex.Count; j++)
                {
                    ProgressBar1.Value = 0;
                    label5_Copy.Content = Convert.ToString(catalogRegex.Count-j-1);
                    UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                    Regex regex = new Regex(@"" + catalogRegex[j].Regex_func + "");
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    int val = 0;
                    excelrange_Name29 = (Excel.Range)xlWorksheet.Cells[i, position_Name9];
                    excelrange_Group29 = (Excel.Range)xlWorksheet.Cells[i, position_Group9];
                    excelrange_Content29 = (Excel.Range)xlWorksheet.Cells[i, position_content9];//Название_Характеристики
                    excelrange_ContentWords29 = (Excel.Range)xlWorksheet.Cells[i, position_content9 + 2];//значення характеристики      
                    string newName = "";
                    string Name = Convert.ToString(excelrange_Name29.Value2);
                    int position_content_new = position_content9;
                    if (Name != null)
                    {
                        MatchCollection matches = regex.Matches(Name);

                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                newName = (match.Value);
                            }
                            s4:
                            if (excelrange_Content29.Value2 != null)
                            {
                                val = position_content_new + 3;
                                excelrange_Content29 = (Excel.Range)xlWorksheet.Cells[i, val];
                                excelrange_ContentWords29 = (Excel.Range)xlWorksheet.Cells[i, val + 2];
                                position_content_new = position_content_new + 3;
                                goto s4;
                            }
                            if (catalogRegex[j].check == true)//вирізати
                            {
                                excelrange_Name29.Value2 = "";
                                excelrange_Name29.Value2 = Name.Replace(newName.Replace("\n", ""), "");
                                excelrange_ContentWords29.Value2 = newName;
                                excelrange_Content29.Value2 = catalogRegex[j].Value;
                            }
                            if (catalogRegex[j].check == false)//залишаємо
                            {
                                excelrange_ContentWords29.Value2 = newName;
                                excelrange_Content29.Value2 = catalogRegex[j].Value;
                            }
                        }
                    }
                }
                }
                #endregion
            label5.Background = Brushes.Green;
            label5_Copy.Background = Brushes.Green;
            //3 характеристики
            try
            {
                int position_Name = 0;//Название_позиции
                int position_Group = 0;//id of group
                int position_content = 0;//id content
                Excel.Range excelrange2;
                Excel.Range excelrange_Name2;
                Excel.Range excelrange_Group2;
                Excel.Range excelrange_Content2;
                Excel.Range excelrange_ContentWords2;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange2 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange2.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }
                    if (Convert.ToString(excelrange2.Value2) == "Идентификатор_группы")
                    {
                        position_Group = i;
                    }
                    if (Convert.ToString(excelrange2.Value2) == "Название_Характеристики")
                    {
                        position_content = i;
                        goto f2;
                    }
                }
                f2:
                int poscont = position_content;
                ProgressBar1.Value = 0;
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name2 = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    excelrange_Group2 = (Excel.Range)xlWorksheet.Cells[i, position_Group];
                    pos1:
                    excelrange_Content2 = (Excel.Range)xlWorksheet.Cells[i, position_content];
                    excelrange_ContentWords2 = (Excel.Range)xlWorksheet.Cells[i, position_content + 2];
                    string Name = Convert.ToString(excelrange_Name2.Value2);
                    foreach (var item in ListOfTypes)
                    {
                        if (excelrange_Content2.Value2 == null)
                        {
                            if (item.GroupId == Convert.ToString(excelrange_Group2.Value2))
                            {
                                if (item.keyWords != null)
                                {
                                    foreach (var el in item.keyWords)
                                    {
                                        if (el.ListValues != null)
                                        {
                                            for(int k=0;k<el.ListValues.Count;k++)
                                           // foreach (var words in el.ListValues)
                                            {
                                                if (Name.Contains(el.ListValues[k].Replace("\n", "")) == true)
                                                {/*!*/
                                                    Name = Name.Replace(el.ListValues[k].Replace("\n", ""), "");
                                                    excelrange_Name2.Value2 = "";
                                                    excelrange_Name2.Value2 = Name;
                                                    excelrange_Content2.Value2 = el.KeyName.Replace("\n", "");//імя характеристики
                                                    if(el.ListValues1[k]=="Value")
                                                    {
                                                        excelrange_ContentWords2.Value2 = el.ListValues[k].Replace("\n", "");//значення характеристики
                                                    }
                                                    else
                                                    {
                                                        excelrange_ContentWords2.Value2 = el.ListValues1[k].Replace("\n", "");//значення характеристики
                                                    }                                                           
                                                    position_content = position_content + 3;
                                                    excelrange_Content2 = (Excel.Range)xlWorksheet.Cells[i, position_content];
                                                    excelrange_ContentWords2 = (Excel.Range)xlWorksheet.Cells[i, position_content + 2];
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                        }
                        else
                        {
                            position_content = position_content + 3;
                            goto pos1;
                        }
                    }
                    position_content = poscont; //start point of position content               
                }
            }
            catch { MessageBox.Show("Помилка, Можливо ви не відкрили файл"); }
            label3.Background = Brushes.Green;
            //усунення помилок ком і крапок
            ProgressBar1.Value = 0;
            ProgressBar1.Maximum = rowCount;
            try
            {
                int position_Name = 0;//Название_позиции
                Excel.Range excelrange2;
                Excel.Range excelrange_Name2;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange2 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange2.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }
                }
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name2 = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    string Name = Convert.ToString(excelrange_Name2.Value2);
                    excelrange_Name2.Value2 = "";
                    excelrange_Name2.Value2 = Name.Replace("\n", "").Replace(",", " ").Replace(".", " ");
                }
            }
            catch { MessageBox.Show("Помилка на кроці вирізання ком і крапок !"); }
            label6.Background = Brushes.Green;
            //
            //4
            try
            {
                int position_Name = 0;//Название_позиции
                int position_Group = 0;//id of group
                int position_content = 0;//id content
                Excel.Range excelrange1;
                Excel.Range excelrange_Name1;
                Excel.Range excelrange_Group1;
                Excel.Range excelrange_Content1;
                Random rand = new Random();
                int val = 0;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange1 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange1.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Идентификатор_группы")
                    {
                        position_Group = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Описание")
                    {
                        position_content = i;
                    }
                }                
                ProgressBar1.Value = 0;
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name1 = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    excelrange_Group1 = (Excel.Range)xlWorksheet.Cells[i, position_Group];
                    excelrange_Content1 = (Excel.Range)xlWorksheet.Cells[i, position_content];

                    foreach (var item in ListOfTypes)
                    {
                        if (Radio1.IsChecked == true)
                        {
                            if (excelrange_Content1.Value2 == null)
                            {
                                if (item.GroupId == Convert.ToString(excelrange_Group1.Value2))
                                {
                                    if (item.Info != null && item.Info.Count != 0)
                                    {
                                        val = rand.Next(0, item.Info.Count);
                                        excelrange_Content1.Value2 = Check_infokey(item.Info[val], Convert.ToString(excelrange_Name1.Value2));
                                    }
                                }
                            }
                        }
                        if (Radio2.IsChecked == true)
                        {
                            if (item.GroupId == Convert.ToString(excelrange_Group1.Value2))
                            {
                                if (item.Info != null && item.Info.Count != 0)
                                {
                                    val = rand.Next(0, item.Info.Count);
                                    excelrange_Content1.Value2 = Check_infokey(item.Info[val], Convert.ToString(excelrange_Name1.Value2));
                                }
                            }
                        }
                    }
                }
            }
            catch { MessageBox.Show("Помилка, Можливо ви не відкрили файл"); }
            label4.Background = Brushes.Green;
            //insert keywords
            ProgressBar1.Value = 0;
            ProgressBar1.Maximum = rowCount;
            XmlSerializer formatter_words = new XmlSerializer(typeof(List<MainWords>));
            try
            {
                using (FileStream fs = new FileStream("KeyWordsCatalog\\ReservCatalogKeyWords.xml", FileMode.Open))
                {
                    List<MainWords> newListWords = (List<MainWords>)formatter_words.Deserialize(fs);
                    MainWords = newListWords;
                }
            }
            catch { MessageBox.Show("Файл ReservCatalogKeyWords.xml відсутній"); }
            try
            {
                Random random = new Random();
                int rnd=0;
                int position_Name_1 = 0;//Название_позиции
                int position_Group_1 = 0;//id of group
                int position_keywords = 0;
                int position_article = 0;
                Excel.Range excelrange1;
                Excel.Range excelrange_Name1;
                Excel.Range excelrange_KeyWords;
                Excel.Range excelrange_Article;
                Excel.Range excelrange_Group1;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange1 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange1.Value2) == "Название_позиции")
                    {
                        position_Name_1 = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Идентификатор_группы")
                    {
                        position_Group_1 = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Ключевые_слова")
                    {
                        position_keywords = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Код_товара")
                    {
                        position_article = i;
                    }
                }
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name1 = (Excel.Range)xlWorksheet.Cells[i, position_Name_1];//Название_позиции
                    excelrange_Group1 = (Excel.Range)xlWorksheet.Cells[i, position_Group_1];//Идентификатор_группы
                    excelrange_KeyWords = (Excel.Range)xlWorksheet.Cells[i, position_keywords];//Ключевые_слова
                    excelrange_Article = (Excel.Range)xlWorksheet.Cells[i, position_article];//Код_товара                   
                    foreach (var item in MainWords)
                    {
                        if (item.GroupId == Convert.ToString(excelrange_Group1.Value2))
                        {
                            if (item.Words != null && item.Words.Count != 0)
                            {
                                rnd = random.Next(0, item.Words.Count);
                                excelrange_KeyWords.Value2 = Check_infoword(item.Words[rnd], Convert.ToString(excelrange_Name1.Value2), Convert.ToString(excelrange_Article.Value2));
                            }
                        }
                    }                   
                }
            }
            catch { MessageBox.Show("Помилка, Можливо ви не відкрили файл"); }            
            label4_Copy.Background = Brushes.Green;
            DateTime timeend = DateTime.Now;
            MessageBox.Show((timeend-timestart).ToString()+" час операції");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);            
        }
        /// <summary>
        /// open regex window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Regex_Click(object sender, RoutedEventArgs e)
        {
            regexWindow = new RegexWindow();
            regexWindow.Show();
        }
        /// <summary>
        /// test regex button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            #region open catalog regex
            List<RegexC> catalogRegex;
            XmlSerializer formatter = new XmlSerializer(typeof(List<RegexC>));
            using (FileStream fs = new FileStream("Catalog\\catalogRegexc.xml", FileMode.Open))
            {
                catalogRegex = (List<RegexC>)formatter.Deserialize(fs);
            }
            #endregion
            #region check and add Измерение_Характеристики            
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;
                int colCount = xlRange.Columns.Count;
                int rowCount = xlRange.Rows.Count;
                try
                {                                      
                    Start123:
                    for (int i = 2; i <= colCount; i++)
                    {
                        excelcells = (Excel.Range)xlWorksheet.Cells[1, i];
                        excelcells1 = (Excel.Range)xlWorksheet.Cells[1, i - 1];
                        if (Convert.ToString(excelcells.Value2) == "Значение_Характеристики" && Convert.ToString(excelcells1.Value2) != "Измерение_Характеристики")//Измерение_Характеристики
                        {
                            Read(excelcells.Address);//func                    
                            goto Start123;// I like use goto ))
                        }
                    }                    
                }
                catch { MessageBox.Show("Виберіть файл для роботи!"); }
            #endregion
            #region set regex values 
           
                    
                    int position_Name = 0;//Название_позиции
                    int position_Group = 0;//id of group
                    int position_content = 0;//id content
                    Excel.Range excelrange2;
                    Excel.Range excelrange_Name2;
                    Excel.Range excelrange_Group2;
                    Excel.Range excelrange_Content2;
                    Excel.Range excelrange_ContentWords2;
                    
                    for (int i = 1; i <= colCount; i++)
                    {
                        excelrange2 = (Excel.Range)xlWorksheet.Cells[1, i];
                        if (Convert.ToString(excelrange2.Value2) == "Название_позиции")
                        {
                            position_Name = i;
                        }
                        if (Convert.ToString(excelrange2.Value2) == "Идентификатор_группы")
                        {
                            position_Group = i;
                        }
                         if (Convert.ToString(excelrange2.Value2) == "Название_Характеристики")
                         {
                            position_content = i;
                    goto f1;
                         }
                    }
                    f1:
            ProgressBar1.Maximum = rowCount;

            for (int j = 0; j < catalogRegex.Count; j++)
                {                
                    ProgressBar1.Value = 0;
                    UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                    Regex regex = new Regex(@"" + catalogRegex[j].Regex_func + "");
                    for (int i = 1; i <= rowCount; i++)
                    {
                        Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty,(double)i});
                        int val = 0;
                        excelrange_Name2 = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                        excelrange_Group2 = (Excel.Range)xlWorksheet.Cells[i, position_Group];
                        excelrange_Content2 = (Excel.Range)xlWorksheet.Cells[i, position_content];//Название_Характеристики
                        excelrange_ContentWords2 = (Excel.Range)xlWorksheet.Cells[i, position_content + 2];//значення характеристики      
                        string newName="";
                        string Name = Convert.ToString(excelrange_Name2.Value2);
                    ///
                        int position_content_new = position_content;
                        MatchCollection matches = regex.Matches(Name);
                        if (matches.Count > 0)
                        {
                            foreach (Match match in matches)
                            {
                                newName = (match.Value);
                            }
                            s4:
                            if (excelrange_Content2.Value2!=null)
                            {
                            val = position_content_new + 3;
                            excelrange_Content2 = (Excel.Range)xlWorksheet.Cells[i, val];
                            excelrange_ContentWords2 = (Excel.Range)xlWorksheet.Cells[i, val + 2];
                            position_content_new = position_content_new + 3;
                            goto s4;
                            }
                            if (catalogRegex[j].check == true)//вирізати
                            {
                                excelrange_Name2.Value2 = "";
                                excelrange_Name2.Value2 = Name.Replace(newName.Replace("\n", ""), "");
                                excelrange_ContentWords2.Value2 = newName;
                                excelrange_Content2.Value2 = catalogRegex[j].Value;
                            }
                            if (catalogRegex[j].check == false)//залишаємо
                            {
                                excelrange_ContentWords2.Value2 = newName;
                                excelrange_Content2.Value2 = catalogRegex[j].Value;
                            }
                        }    
                    }
                }
           
            #endregion
            MessageBox.Show("Yesss!");
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);           

        }

        private void Koma_Click(object sender, RoutedEventArgs e)
        {
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;
            ProgressBar1.Value = 0;
            ProgressBar1.Maximum = rowCount;
            try
            {
                int position_Name = 0;//Название_позиции
                Excel.Range excelrange2;
                Excel.Range excelrange_Name2;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange2 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange2.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }

                }
                UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name2 = (Excel.Range)xlWorksheet.Cells[i, position_Name];
                    string Name = Convert.ToString(excelrange_Name2.Value2);       
                    excelrange_Name2.Value2 = "";
                    excelrange_Name2.Value2 = Name.Replace("\n", "").Replace(","," ").Replace("."," ");                    
                }

            }
            catch { MessageBox.Show("Помилка на кроці вирізання ком і крапок !"); }
            label6.Background = Brushes.Green;
        }

        private void Button_KeyWords_Click(object sender, RoutedEventArgs e)
        {
            KeyWordsWindow = new KeyWordsWindow();
            KeyWordsWindow.Show();
        }
/// <summary>
/// генеруються ключові слова з каталогу
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
        private void Button_KeyWords1_Click(object sender, RoutedEventArgs e)
        {
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            ProgressBar1.Value = 0;
            UpdateProgressBarDelegate updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar1.SetValue);
            int colCount = xlRange.Columns.Count;
            int rowCount = xlRange.Rows.Count;

            ProgressBar1.Maximum = rowCount;
            XmlSerializer formatter = new XmlSerializer(typeof(List<MainWords>));
           
                using (FileStream fs = new FileStream("KeyWordsCatalog\\ReservCatalogKeyWords.xml", FileMode.Open))
                {
                    List<MainWords> newListWords = (List<MainWords>)formatter.Deserialize(fs);
                    MainWords = newListWords;
                }
                int rnd = 0;
                Random random = new Random();
                int position_Name = 0;//Название_позиции
                int position_Group = 0;//id of group
                int position_keywords = 0;
                int position_article = 0;
                Excel.Range excelrange1;
                Excel.Range excelrange_Name1;
                Excel.Range excelrange_KeyWords;
                Excel.Range excelrange_Article;
                Excel.Range excelrange_Group1;
                for (int i = 1; i <= colCount; i++)
                {
                    excelrange1 = (Excel.Range)xlWorksheet.Cells[1, i];
                    if (Convert.ToString(excelrange1.Value2) == "Название_позиции")
                    {
                        position_Name = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Идентификатор_группы")
                    {
                        position_Group = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Ключевые_слова")
                    {
                        position_keywords = i;
                    }
                    if (Convert.ToString(excelrange1.Value2) == "Код_товара")
                    {
                        position_article = i;
                    }
                }
                for (int i = 1; i <= rowCount; i++)
                {
                    Dispatcher.Invoke(updatePbDelegate, System.Windows.Threading.DispatcherPriority.Background, new object[] { ProgressBar.ValueProperty, (double)i });
                    excelrange_Name1 = (Excel.Range)xlWorksheet.Cells[i, position_Name];//Название_позиции
                    excelrange_Group1 = (Excel.Range)xlWorksheet.Cells[i, position_Group];//Идентификатор_группы
                    excelrange_KeyWords = (Excel.Range)xlWorksheet.Cells[i, position_keywords];//Ключевые_слова
                    excelrange_Article = (Excel.Range)xlWorksheet.Cells[i, position_article];//Код_товара
                    foreach (var item in MainWords)
                    {
                        if (item.GroupId == Convert.ToString(excelrange_Group1.Value2))
                        {
                            if (item.Words != null && item.Words.Count!=0)
                            {
                                rnd = random.Next(0, item.Words.Count);                           
                                excelrange_KeyWords.Value2 = Check_infoword(item.Words[rnd], Convert.ToString(excelrange_Name1.Value2), Convert.ToString(excelrange_Article.Value2));                            
                            }
                        }
                    }
                }
        }

        private void Button_Restruct_Click(object sender, RoutedEventArgs e)
        {
            restructWindow = new RestructWindow();
            restructWindow.Show();
        }

        private void ButtonRestruct_Click(object sender, RoutedEventArgs e)
        {
            reStruct = new ReStruct();
            XmlSerializer formatter = new XmlSerializer(typeof(ReStruct));
            try
            {
                using (FileStream fs = new FileStream("Restruct\\Restruct.xml", FileMode.Open))
                {
                    ReStruct newrestruct = (ReStruct)formatter.Deserialize(fs);
                    reStruct = newrestruct;
                }
            }
            catch { MessageBox.Show("Відсутній файл Restruct.xml"); }
        }
    }
}