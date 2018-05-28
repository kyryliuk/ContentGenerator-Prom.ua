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
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        ListGroupNames.ItemsSource = item.Info;
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
                        if (item.Info[ListGroupNames.SelectedIndex].ToString() != null)
                        {
                            item.Info[ListGroupNames.SelectedIndex] = EditWindow.infocontent;
                            ListGroupNames.ItemsSource = null;
                            ListGroupNames.ItemsSource = item.Info;
                        }
                    }
                }
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }

        private void ListGroupNames_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeView.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListGroupNames.SelectedIndex].ToString() != null)
                        {
                            editWindowC = new EditWindow(item.Info[ListGroupNames.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindowC.Show();
                            this.IsEnabled = false;
                            editWindowC.Closed += ClosedEdit;
                        }
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }
    }
}
