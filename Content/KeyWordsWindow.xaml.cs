using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Content
{
    /// <summary>
    /// Interaction logic for KeyWordsWindow.xaml
    /// </summary>
    public partial class KeyWordsWindow : Window
    {
        EditWindow editWindow;
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        Excel.Range excelcells1;
        Excel.Range excelcellsGroup;
        Excel.Range excelcellsParent;
        public string PathCatalog = "KeyWordsCatalog\\ReservCatalogKeyWords.xml";
        List<MainWords> MainWords;
        List<MainWords> MainWords_catalog;
        string pathfile;
        List<TreeViewItem> treeViewItems;
        List<TreeViewItem> treeviewitemcatalog;
        public KeyWordsWindow()
        {
            InitializeComponent();
            UpdateList();
        }
        public void UpdateList()
        {
            ListFiles.Items.Clear();
            DirectoryInfo d = new DirectoryInfo("KeyWordsCatalog");//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xml"); //Getting Text files           
            foreach (FileInfo file in Files)
            {
                ListFiles.Items.Add(file.Name);
            }
        }

        private void ListFiles_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TreeView.Items.Clear();
            ListViewWords.ItemsSource = null;
            ListViewWords.Items.Clear();
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<MainWords>));
                try
                {
                    using (FileStream fs = new FileStream("KeyWordsCatalog\\" + ListFiles.SelectedItem.ToString() + "", FileMode.Open))
                    {
                        List<MainWords> newListWords = (List<MainWords>)formatter.Deserialize(fs);
                        MainWords = newListWords;
                    }
                }
                catch { MessageBox.Show("Файл ReservCatalogKeyWords.xml відсутній"); }
                treeviewitemcatalog = new List<TreeViewItem>();
                foreach (var item in MainWords)//create header for tree
                {
                    if (item.ParentId == null)
                    {
                        treeviewitemcatalog.Add(new TreeViewItem() { Header = item.value });
                    }
                }
                //
                foreach (var item in MainWords)//build tree
                {
                    if (item.ParentId != null)
                    {
                        foreach (var grid in MainWords)
                        {
                            if (item.ParentId == grid.GroupId)
                            {
                                foreach (var ite1 in treeviewitemcatalog)
                                    if (grid.value == Convert.ToString(ite1.Header))
                                    {
                                        ite1.Items.Add(new TreeViewItem() { Header = item.value });
                                    }
                            }
                        }
                    }
                }
                foreach (var it in treeviewitemcatalog) { TreeView.Items.Add(it); }
            }
            catch { MessageBox.Show("Виберіть файл"); }
        }

        private void TreeView_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ListViewWords.ItemsSource = null;
            ListViewWords.Items.Clear();
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {                         
                        ListViewWords.ItemsSource = item.Words;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }

        private void Button_Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<MainWords>));
                string reservname = "KeyWordsCatalog\\ReservCatalogKeyWords" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
                if (MainWords != null)
                {
                    if (File.Exists("KeyWordsCatalog\\ReservCatalogKeyWords.xml"))//temp delete for real view info
                    {
                        File.Delete("KeyWordsCatalog\\ReservCatalogKeyWords.xml");
                    }
                    using (FileStream fs = new FileStream(reservname, FileMode.Create))
                    {
                        formatter.Serialize(fs, MainWords);
                    }
                    using (FileStream fs = new FileStream("KeyWordsCatalog\\ReservCatalogKeyWords.xml", FileMode.OpenOrCreate))
                    {
                        formatter.Serialize(fs, MainWords);
                    }
                    MessageBox.Show("Резервна копія каталога збережена\n " + reservname + "");
                }
                else { MessageBox.Show("Ви не відкрили каталог для роботи!"); }
            }
            catch { MessageBox.Show("Помилка"); }
        }
        /// <summary>
        /// update data!
        /// </summary>
        public void UpdateData()
        {
            MainWords_catalog = null;
            TreeView.Items.Clear();
            ListViewWords.ItemsSource = null;
            ListViewWords.Items.Clear();
            XmlSerializer formatter = new XmlSerializer(typeof(List<MainWords>));
            try
            {
                using (FileStream fs = new FileStream("KeyWordsCatalog\\ReservCatalogKeyWords.xml", FileMode.Open))
                {
                    List<MainWords> newListWords = (List<MainWords>)formatter.Deserialize(fs);
                    MainWords_catalog = newListWords;
                }
            }
            catch { MessageBox.Show("Файл ReservCatalogKeyWords.xml відсутній"); }
            foreach (var listreal in MainWords)
            {
                foreach (var listfile in MainWords_catalog)
                {
                    if (listreal.value == listfile.value)
                    {
                        listreal.Words = listfile.Words;
                    }
                }
            }
            //
            treeViewItems = new List<TreeViewItem>();
            foreach (var item in MainWords)//create header for tree
            {
                if (item.ParentId == null)
                {
                    treeViewItems.Add(new TreeViewItem() { Header = item.value });
                }
            }
            //
            foreach (var item in MainWords)//build tree
            {
                if (item.ParentId != null)
                {
                    foreach (var grid in MainWords)
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

            foreach (var it in treeViewItems) { TreeView.Items.Add(it); }
        }

        private void Button_Open_Click(object sender, RoutedEventArgs e)
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
                MainWords = null;
                MainWords = new List<MainWords>();
                for (int j = 2; j <= colRows; j++)
                {
                    excelcells1 = (Excel.Range)xlWorksheet.Cells[j, 2];
                    excelcellsGroup = (Excel.Range)xlWorksheet.Cells[j, 3];
                    excelcellsParent = (Excel.Range)xlWorksheet.Cells[j, 5];
                    if (Convert.ToString(excelcells1.Value2) != null)
                    {
                        MainWords.Add(new MainWords(Convert.ToString(excelcells1.Value2), Convert.ToString(excelcellsGroup.Value2), Convert.ToString(excelcellsParent.Value2)));
                    }
                }
                //XmlSerializer formatter = new XmlSerializer(typeof(List<MainWords>));
                //using (FileStream fs = new FileStream("KeyWordsCatalog\\ReservCatalogKeyWords.xml", FileMode.OpenOrCreate))
                //{
                //    formatter.Serialize(fs, MainWords);
                //}
                UpdateData();
            }
            catch { MessageBox.Show("Виберіть файл"); }
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                ListViewWords.ItemsSource = null;
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {
                        item.Words.Add("keyword");
                        ListViewWords.ItemsSource = item.Words;
                    }
                }
            }
            catch { MessageBox.Show("Виберіть файл для роботи!"); }
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {
                        item.Words.RemoveAt(ListViewWords.SelectedIndex);
                        ListViewWords.ItemsSource = null;
                        ListViewWords.ItemsSource = item.Words;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент"); }
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {
                        if (item.Words[ListViewWords.SelectedIndex].ToString() != null)
                        {
                            editWindow = new EditWindow(item.Words[ListViewWords.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindow.Show();
                            this.IsEnabled = false;
                            editWindow.Closed += ClosedEdit;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        public void ClosedEdit(object eventt, System.EventArgs eventArgs)//delegate for event
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {
                        if (item.Words[ListViewWords.SelectedIndex].ToString() != null)
                        {
                            item.Words[ListViewWords.SelectedIndex] = EditWindow.infocontent;
                            ListViewWords.ItemsSource = null;
                            ListViewWords.ItemsSource = item.Words;
                        }
                    }
                }
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }

        private void ListViewWords_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in MainWords)
                {
                    if (selit == item.value)
                    {
                        if (item.Words[ListViewWords.SelectedIndex].ToString() != null)
                        {
                            editWindow = new EditWindow(item.Words[ListViewWords.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindow.Show();
                            this.IsEnabled = false;
                            editWindow.Closed += ClosedEdit;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
    }
}
