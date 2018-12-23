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
    /// Interaction logic for RestructWindow.xaml
    /// </summary>
    public partial class RestructWindow : Window
    {
        ReStruct reStruct;
        public RestructWindow()
        {
            ReadXML();
            
            
            //reStruct = new ReStruct
            //{
            //    Product_Del = new List<string>(),
            //    Group_Del = new List<string>(),
            //    Group_Edit = new List<twolist>(),
            //    Product_Redirect = new List<twolist>(),
            //    Group_Redirect = new List<twolist>(),
            //    Group_New = new List<threelist>()
            //};
            InitializeComponent();
            List_product_del.ItemsSource = reStruct.Product_Del;
            List_group_del.ItemsSource = reStruct.Group_Del;
            List_group_edit.ItemsSource = reStruct.Group_Edit;
            List_product_redirect.ItemsSource = reStruct.Product_Redirect;
            List_group_redirect.ItemsSource = reStruct.Group_Redirect;
            List_group_create.ItemsSource = reStruct.Group_New;
        }
        public void ReadXML()
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
        //add
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            List_product_del.ItemsSource = null;
            List_product_del.Items.Clear();
            reStruct.Product_Del.Add(TextBox_product_del.Text);
            List_product_del.ItemsSource = reStruct.Product_Del;
            TextBox_product_del.Text = "";
        }
        //edit
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TextBox_product_del.Text = List_product_del.SelectedItem.ToString();
        }
        //remove
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {            
            reStruct.Product_Del.RemoveAt(List_product_del.SelectedIndex);
            List_product_del.ItemsSource = null;
            List_product_del.ItemsSource = reStruct.Product_Del;

        }

        private void SAVe_Click(object sender, RoutedEventArgs e)
        {
           
                XmlSerializer formatter = new XmlSerializer(typeof(ReStruct));
                string reservname = "Restruct\\ReservRestruct" + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss") + ".xml";
                if (reStruct != null)
                {
                    using (FileStream fs = new FileStream(reservname, FileMode.Create))
                    {
                        formatter.Serialize(fs, reStruct);
                    }
                    if (File.Exists("Restruct\\Restruct.xml"))//temp delete for real view info
                    {
                        File.Delete("Restruct\\Restruct.xml");
                    }
                    using (FileStream fs = new FileStream("Restruct\\Restruct.xml", FileMode.OpenOrCreate))
                    {
                        formatter.Serialize(fs, reStruct);
                    }
                    MessageBox.Show("Резервна копія каталога збережена\n " + reservname + "");
                }
                else { MessageBox.Show("Ви не відкрили каталог для роботи!"); }
           
        }
        //add
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            List_group_del.ItemsSource = null;
            List_group_del.Items.Clear();
            reStruct.Group_Del.Add(TextBoxGroupDel.Text);
            List_group_del.ItemsSource = reStruct.Group_Del;
            TextBox_product_del.Text = "";
        }
        //remove
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            reStruct.Group_Del.RemoveAt(List_group_del.SelectedIndex);
            List_group_del.ItemsSource = null;
            List_group_del.ItemsSource = reStruct.Product_Del;
        }
        //add
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            List_group_edit.ItemsSource = null;
            List_group_edit.Items.Clear();
            reStruct.Group_Edit.Add(new twolist( TextBox_edit_group_old.Text,TextBox_edit_group_new.Text));
            List_group_edit.ItemsSource = reStruct.Group_Edit;
            TextBox_edit_group_old.Text = "";
            TextBox_edit_group_new.Text = "";
        }       
        //remove
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            reStruct.Group_Edit.RemoveAt(List_group_edit.SelectedIndex);
            List_group_edit.ItemsSource = null;
            List_group_edit.ItemsSource = reStruct.Group_Edit;
        }
        //add
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            List_product_redirect.ItemsSource = null;
            List_product_redirect.Items.Clear();
            reStruct.Product_Redirect.Add(new twolist(TextBox_product_redirect_name.Text, TextBox_product_redirect_groupID.Text));
            List_product_redirect.ItemsSource = reStruct.Product_Redirect;
            TextBox_product_redirect_name.Text = "";
            TextBox_product_redirect_groupID.Text = "";
        }
        //remove
        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            reStruct.Product_Redirect.RemoveAt(List_group_edit.SelectedIndex);
            List_group_edit.ItemsSource = null;
            List_group_edit.ItemsSource = reStruct.Product_Redirect;
        }
        //add
        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            List_group_redirect.ItemsSource = null;
            List_group_redirect.Items.Clear();
            reStruct.Group_Redirect.Add(new twolist(TextBox_Group_redirect_oldID.Text, TextBox_Group_redirect_NewID.Text));
            List_group_redirect.ItemsSource = reStruct.Group_Redirect;
            TextBox_Group_redirect_oldID.Text = "";
            TextBox_Group_redirect_NewID.Text = "";
        }
        //remove
        private void Button_Click_10(object sender, RoutedEventArgs e)
        {
            reStruct.Group_Redirect.RemoveAt(List_group_redirect.SelectedIndex);
            List_group_redirect.ItemsSource = null;
            List_group_redirect.ItemsSource = reStruct.Group_Redirect;
        }

        private void Button_Click_11(object sender, RoutedEventArgs e)
        {
            List_group_create.ItemsSource = null;
            List_group_create.Items.Clear();
            reStruct.Group_New.Add(new  threelist(TextBox_newgroup_name.Text, TextBox_newgroup_ID.Text, TextBox_newgroup_ParentID.Text));
            List_group_create.ItemsSource = reStruct.Group_New;
            TextBox_newgroup_name.Text = "";
            TextBox_newgroup_ID.Text = "";
            TextBox_newgroup_ParentID.Text = "";
        }

        private void Button_Click_12(object sender, RoutedEventArgs e)
        {
            reStruct.Group_New.RemoveAt(List_group_create.SelectedIndex);
            List_group_create.ItemsSource = null;
            List_group_create.ItemsSource = reStruct.Group_New;
        }
    }
}
