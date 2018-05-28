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
    /// Interaction logic for CatalogWindow.xaml
    /// </summary>
    public partial class CatalogWindow : Window
    {
        public CatalogWindow()
        {
            InitializeComponent();
            UpdateList();
        }

        private void ButtonUse_Click(object sender, RoutedEventArgs e)
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
            catch { MessageBox.Show("Виберіть файл"); }
        }

        private void Button_delete_catalog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists("Catalog\\" + ListCatalog.SelectedItem.ToString() + ""))//temp delete for real view info
                {
                    File.Delete("Catalog\\" + ListCatalog.SelectedItem.ToString() + "");
                }
                UpdateList();
            }
            catch { MessageBox.Show("Виберіть файл"); }
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
    }
}
