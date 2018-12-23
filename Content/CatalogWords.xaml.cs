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

namespace Content
{
    /// <summary>
    /// Interaction logic for CatalogWords.xaml
    /// </summary>
    public partial class CatalogWords : Window
    {
        List<ExcelTypes> catalogsfromreserv;
        List<TreeViewItem> treeviewitemcatalog;
        EditWindow editWindowC;
        public CatalogWords()
        {
            InitializeComponent();
            UpdateList();
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {

        }
        public void UpdateList()
        {
            ListFiles.Items.Clear();
            DirectoryInfo d = new DirectoryInfo("Catalog");//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xml"); //Getting Text files           
            foreach (FileInfo file in Files)
            {
                ListFiles.Items.Add(file.Name);
            }
        }

        private void ListFiles_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TreeView.Items.Clear();
            ListGroupNames.ItemsSource = null;
            ListGroupNames.Items.Clear();
            ListKeyWords.ItemsSource = null;
            ListKeyWords.Items.Clear();
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
                try
                {
                    using (FileStream fs = new FileStream("Catalog\\" + ListFiles.SelectedItem.ToString() + "", FileMode.Open))
                    {
                        List<ExcelTypes> newListWorkers = (List<ExcelTypes>)formatter.Deserialize(fs);
                        catalogsfromreserv = newListWorkers;
                    }
                }
                catch { MessageBox.Show("Файл CatalogContent.xml відсутній"); }
                treeviewitemcatalog = new List<TreeViewItem>();
                foreach (var item in catalogsfromreserv)//create header for tree
                {
                    if (item.ParentId == null)
                    {
                        treeviewitemcatalog.Add(new TreeViewItem() { Header = item.value });
                    }
                }
                //
                foreach (var item in catalogsfromreserv)//build tree
                {
                    if (item.ParentId != null)
                    {
                        foreach (var grid in catalogsfromreserv)
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
            ListKeyWords.ItemsSource = null;
            ListKeyWords.Items.Clear();
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    List<string> vs = new List<string>();
                    if (selit == item.value)
                    {
                        foreach (var el in item.keyWords)
                        { vs.Add(el.KeyName); }
                        ListGroupNames.ItemsSource = vs;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }       

        private void ListGroupNames_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            //try
            //{
            //    TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
            //    string selit = items.Header.ToString();
            //    foreach (var item in catalogsfromreserv)
            //    {
            //        if (selit == item.value)
            //        {
            //            if (item.Info[ListGroupNames.SelectedIndex].ToString() != null)
            //            {
            //                editWindowC = new EditWindow(item.Info[ListGroupNames.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
            //                editWindowC.Show();
            //                this.IsEnabled = false;
            //                editWindowC.Closed += ClosedEdit;
            //            }
            //        }
            //    }
            //}
            //catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }

        private void Button_Add1_Click(object sender, RoutedEventArgs e)
        {
            
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                ListGroupNames.ItemsSource = null;
                ListKeyWords.ItemsSource = null;
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        item.keyWords.Add(new KeyWords("Keyword",new List<string>{ "Marketing"},new List<string> { "Value"}));
                        List<string> lstr = new List<string>();
                        List<twolist> twolists = new List<twolist>();
                        foreach (var el in item.keyWords)
                        {
                            lstr.Add(el.KeyName);                         
                            for (int i = 0; i < el.ListValues.Count; i++)
                            {
                             twolists.Add(new twolist(el.ListValues[i], el.ListValues1[i]));                            
                            }
                        }
                        ListGroupNames.ItemsSource = null;
                        ListKeyWords.ItemsSource = null;
                        ListGroupNames.Items.Clear();
                        ListKeyWords.Items.Clear();
                        ListGroupNames.ItemsSource = lstr;
                        /**/
                        ListKeyWords.ItemsSource = twolists;
                    }
                }
               
            
        }

        private void Button_delete1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        item.keyWords.RemoveAt(ListGroupNames.SelectedIndex);                    
                    List<string> lstr = new List<string>();
                    foreach (var el in item.keyWords) { lstr.Add(el.KeyName); }
                    ListGroupNames.ItemsSource = null;
                    ListKeyWords.ItemsSource = null;
                    ListGroupNames.Items.Clear();
                    ListKeyWords.Items.Clear();
                    ListGroupNames.ItemsSource = lstr;

                }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент"); }
        }

        private void ListGroupNames_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ListKeyWords.ItemsSource = null;
            ListKeyWords.Items.Clear();
            List<twolist> twolists = new List<twolist>();
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        for (int i = 0; i < item.keyWords[ListGroupNames.SelectedIndex].ListValues.Count; i++)
                        {
                            twolists.Add(new twolist(item.keyWords[ListGroupNames.SelectedIndex].ListValues[i], item.keyWords[ListGroupNames.SelectedIndex].ListValues1[i]));
                        }
                      ListKeyWords.ItemsSource=twolists;                           
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент"); }
        }

        private void Button_Add2_Click(object sender, RoutedEventArgs e)
        {
            ListKeyWords.ItemsSource = null;
            ListKeyWords.Items.Clear();
            List<twolist> twolists = new List<twolist>();
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                ListKeyWords.ItemsSource = null;
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        item.keyWords[ListGroupNames.SelectedIndex].ListValues.Add("Marketing");
                        item.keyWords[ListGroupNames.SelectedIndex].ListValues1.Add("Value");
                        for (int i = 0; i < item.keyWords[ListGroupNames.SelectedIndex].ListValues.Count; i++)
                        {
                            twolists.Add(new twolist(item.keyWords[ListGroupNames.SelectedIndex].ListValues[i], item.keyWords[ListGroupNames.SelectedIndex].ListValues1[i]));
                        }
                        ListKeyWords.ItemsSource = twolists;                        
                    }
                }
            }
            catch { MessageBox.Show("Виберіть елемент для роботи! Або добавте новий"); }
        }

        private void Button_edit1_Click(object sender, RoutedEventArgs e)
        {            
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.keyWords[ListGroupNames.SelectedIndex].KeyName!= null)
                        {
                            editWindowC = new EditWindow(item.keyWords[ListGroupNames.SelectedIndex].KeyName);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindowC.Show();
                            this.IsEnabled = false;
                            editWindowC.Closed += ClosedEdit;
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
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.keyWords[ListGroupNames.SelectedIndex].KeyName != null)
                        {
                            item.keyWords[ListGroupNames.SelectedIndex].KeyName = EditWindow.infocontent;      //! 
                        List<string> lstr = new List<string>();
                        foreach (var el in item.keyWords)
                        { lstr.Add(el.KeyName); }                  
                        ListGroupNames.ItemsSource = null;
                        ListGroupNames.Items.Clear();
                        ListGroupNames.ItemsSource = lstr;
                    }
                    }
                }
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }

        private void Button_delete2_Click(object sender, RoutedEventArgs e)
        {
            List<twolist> twolists = new List<twolist>();
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        item.keyWords[ListGroupNames.SelectedIndex].ListValues.RemoveAt(ListKeyWords.SelectedIndex);                                               
                        ListKeyWords.ItemsSource = null;                        
                        ListKeyWords.Items.Clear();
                        for (int i = 0; i < item.keyWords[ListGroupNames.SelectedIndex].ListValues.Count; i++)
                        {
                            twolists.Add(new twolist (item.keyWords[ListGroupNames.SelectedIndex].ListValues[i], item.keyWords[ListGroupNames.SelectedIndex].ListValues1[i]));
                        }
                        ListKeyWords.ItemsSource = twolists;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент"); }
        }

        private void Button_edit2_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.keyWords[ListGroupNames.SelectedIndex].ListValues[ListKeyWords.SelectedIndex] != null)
                        {
                            editWindowC = new EditWindow(item.keyWords[ListGroupNames.SelectedIndex].ListValues[ListKeyWords.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindowC.Show();
                            this.IsEnabled = false;
                            editWindowC.Closed += ClosedEditKey;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        public void ClosedEditKey(object eventt, System.EventArgs eventArgs)//delegate for event
        {
            List<twolist> twolists = new List<twolist>();
            //  try
            // {
            TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.keyWords[ListGroupNames.SelectedIndex].ListValues[ListKeyWords.SelectedIndex] != null)
                        {
                            item.keyWords[ListGroupNames.SelectedIndex].ListValues[ListKeyWords.SelectedIndex] = EditWindow.infocontent;                              
                            ListKeyWords.ItemsSource = null;
                            ListKeyWords.Items.Clear();
                        for (int i = 0; i < item.keyWords[ListGroupNames.SelectedIndex].ListValues.Count; i++)
                        {
                            twolists.Add(new twolist(item.keyWords[ListGroupNames.SelectedIndex].ListValues[i], item.keyWords[ListGroupNames.SelectedIndex].ListValues1[i]));
                        }
                            ListKeyWords.ItemsSource = twolists;
                        }
                    }
                }
                this.IsEnabled = true;
           // }
           // catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
                string reservname = "Catalog\\ReservCatalogContent" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
                if (catalogsfromreserv != null)
                {
                    if (File.Exists("Catalog\\CatalogContent.xml"))//temp delete for real view info
                    {
                        File.Delete("Catalog\\CatalogContent.xml");
                    }
                    using (FileStream fs = new FileStream(reservname, FileMode.Create))
                    {
                        formatter.Serialize(fs, catalogsfromreserv);
                    }
                   
                    using (FileStream fs = new FileStream("Catalog\\CatalogContent.xml", FileMode.OpenOrCreate))
                    {
                        formatter.Serialize(fs, catalogsfromreserv);
                    }
                    MessageBox.Show("Резервна копія каталога збережена\n " + reservname + "");
                }
                else { MessageBox.Show("Ви не відкрили каталог для роботи!"); }
            }
            catch { MessageBox.Show("Помилка"); }
        }

        private void ListFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void UseInMain_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.PathCatalog = "Catalog\\" + ListFiles.SelectedItem.ToString() + "";
                this.Close();
            }
            catch { MessageBox.Show("Виберіть елемент"); }
        }
        /// <summary>
        /// кнопка редагувати значення 
        /// це є з 2-ї колонки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {            
           try
            {
                var selectedItem = (dynamic)ListKeyWords.SelectedItems[0];
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.keyWords[ListGroupNames.SelectedIndex].ListValues1[ListKeyWords.SelectedIndex] != null)
                        {
                            editWindowC = new EditWindow(selectedItem.ListValues1);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindowC.Show();
                            this.IsEnabled = false;
                            editWindowC.Closed += ClosedEditValue;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
        public void ClosedEditValue(object eventt, System.EventArgs eventArgs)//delegate for event
        {
            List<twolist> twolists = new List<twolist>();
            //  try
            // {
            TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
            string selit = items.Header.ToString();
            foreach (var item in catalogsfromreserv)
            {
                if (selit == item.value)
                {
                    if (item.keyWords[ListGroupNames.SelectedIndex].ListValues1[ListKeyWords.SelectedIndex] != null)
                    {
                        item.keyWords[ListGroupNames.SelectedIndex].ListValues1[ListKeyWords.SelectedIndex] = EditWindow.infocontent;
                        ListKeyWords.ItemsSource = null;
                        ListKeyWords.Items.Clear();
                        for (int i = 0; i < item.keyWords[ListGroupNames.SelectedIndex].ListValues.Count; i++)
                        {
                            twolists.Add(new twolist(item.keyWords[ListGroupNames.SelectedIndex].ListValues[i], item.keyWords[ListGroupNames.SelectedIndex].ListValues1[i]));
                        }
                        ListKeyWords.ItemsSource = twolists;
                    }
                }
            }
            this.IsEnabled = true;
            // }
            // catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }
    }
}
