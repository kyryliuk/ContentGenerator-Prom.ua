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

namespace Content
{
    /// <summary>
    /// Interaction logic for RegexWindow.xaml
    /// </summary>
    public partial class RegexWindow : System.Windows.Window
    {
        TestRegexWindow testRegexWindow;
     public static  List<RegexC> catalogRegexc;
        public RegexWindow()
        {
            UpdateList();
            InitializeComponent();
            RegexView.ItemsSource = null;
            RegexView.Items.Clear();
            RegexView.ItemsSource = catalogRegexc;
        }
        public void UpdateList()
        {
            List<RegexC> newListWorkers;
            XmlSerializer formatter = new XmlSerializer(typeof(List<RegexC>));
            
                using (FileStream fs = new FileStream("Catalog\\catalogRegexc.xml", FileMode.Open))
                {
                    newListWorkers = (List<RegexC>)formatter.Deserialize(fs);                
                }
            catalogRegexc = new List<RegexC>();
            foreach (var item in newListWorkers)
            {
                catalogRegexc.Add(new RegexC(item.Regex_func,item.Value,item.check));
            }
          
        }
        /// <summary>
        /// run
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
                   
        }
        /// <summary>
        /// save
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(List<RegexC>));                
                if (catalogRegexc != null)
                {
                    if (File.Exists("Catalog\\catalogRegexc.xml"))//temp delete for real view info
                    {
                        File.Delete("Catalog\\catalogRegexc.xml");
                    }
                    using (FileStream fs = new FileStream("Catalog\\catalogRegexc.xml", FileMode.Create))
                    {
                        formatter.Serialize(fs, catalogRegexc);
                    }
                }
                else { MessageBox.Show("Ви не відкрили каталог для роботи!"); }
            }
            catch { MessageBox.Show("Помилка"); }
        }
        /// <summary>
        /// add
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            RegexView.ItemsSource = null;
            RegexView.Items.Clear();
            catalogRegexc.Add(new RegexC(TextBox_input_regex.Text,TextBox_value.Text,Check_Box1.IsChecked.Value));
            RegexView.ItemsSource = catalogRegexc;
        }
        /// <summary>
        /// delete
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            catalogRegexc.RemoveAt(RegexView.SelectedIndex);
            RegexView.ItemsSource = null;
            RegexView.Items.Clear();
            RegexView.ItemsSource = catalogRegexc;
        }

        private void Button_Edit_Click(object sender, RoutedEventArgs e)
        {
            TextBox_value.Text = "";
            TextBox_input_regex.Text = "";
            TextBox_input_regex.Text = catalogRegexc[RegexView.SelectedIndex].Regex_func;
            TextBox_value.Text = catalogRegexc[RegexView.SelectedIndex].Value;
            Check_Box1.IsChecked = catalogRegexc[RegexView.SelectedIndex].check;
        }

        private void Button_test_Click(object sender, RoutedEventArgs e)
        {
            testRegexWindow = new TestRegexWindow();
            testRegexWindow.Show();
        }
    }
}
