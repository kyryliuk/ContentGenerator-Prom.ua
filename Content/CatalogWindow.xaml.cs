﻿using System;
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
    /// Interaction logic for CatalogWindow.xaml
    /// </summary>
    public partial class CatalogWindow : Window
    {
        List<ExcelTypes> catalogsfromreserv;
        List<TreeViewItem> treeviewitemcatalog;
        EditWindow editWindowC;
        
        public CatalogWindow()
        {
            InitializeComponent();
            UpdateList();
        }
       
        public void ButtonUse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainWindow.PathCatalog = "Catalog\\" + ListCatalog.SelectedItem.ToString() + "";
                this.Close();
            }
            catch { MessageBox.Show("Виберіть елемент"); }
        }

        private void Button_delete_catalog_Click(object sender, RoutedEventArgs e)
        {
            TreeViewCatalog.Items.Clear();
            try
            {
                if (File.Exists("Catalog\\" + ListCatalog.SelectedItem.ToString() + ""))//temp delete for real view info
                {
                    File.Delete("Catalog\\" + ListCatalog.SelectedItem.ToString() + "");
                }
                UpdateList();
            }
            catch { MessageBox.Show("Виберіть елемент зі списку"); }
        }
        public void UpdateList()
        {
            ListCatalog.Items.Clear();
            DirectoryInfo d = new DirectoryInfo("Catalog");//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xml"); //Getting Text files           
            foreach (FileInfo file in Files)
            {
                ListCatalog.Items.Add(file.Name);
            }
        }

        private void TreeViewCatalog_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeViewCatalog.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        ListBoxCatalog.ItemsSource = item.Info;
                    }
                }
            }
            catch { MessageBox.Show("Ви не вибрали елемент!"); }
        }

        private void ListCatalog_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TreeViewCatalog.Items.Clear();
            ListBoxCatalog.ItemsSource = null;
            ListBoxCatalog.Items.Clear();
            try
            {              
                XmlSerializer formatter = new XmlSerializer(typeof(List<ExcelTypes>));
                try
                {
                    using (FileStream fs = new FileStream("Catalog\\" + ListCatalog.SelectedItem.ToString() + "", FileMode.Open))
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
                foreach (var it in treeviewitemcatalog) { TreeViewCatalog.Items.Add(it); }
            }
            catch { MessageBox.Show("Виберіть файл"); }
        }

        private void ListBoxCatalog_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                TreeViewItem items = TreeViewCatalog.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListBoxCatalog.SelectedIndex].ToString() != null)
                        {
                            editWindowC = new EditWindow(item.Info[ListBoxCatalog.SelectedIndex]);//async await !!!!!!!!!!!!!!!!!!!!!!!!!!!!!                              
                            editWindowC.Show();
                            this.IsEnabled = false;
                            editWindowC.Closed +=ClosedEdit;
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
                TreeViewItem items = TreeViewCatalog.SelectedItem as TreeViewItem;
                string selit = items.Header.ToString();
                foreach (var item in catalogsfromreserv)
                {
                    if (selit == item.value)
                    {
                        if (item.Info[ListBoxCatalog.SelectedIndex].ToString() != null)
                        {
                            item.Info[ListBoxCatalog.SelectedIndex] = EditWindow.infocontent;
                            ListBoxCatalog.ItemsSource = null;
                            ListBoxCatalog.ItemsSource = item.Info;
                        }
                    }
                }
                this.IsEnabled = true;
            }
            catch { MessageBox.Show("Помилка"); this.IsEnabled = true; }
        }
        protected override void OnClosed(EventArgs e) => base.OnClosed(e);
    }
}
