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
using System.Windows.Shapes;

namespace Content
{
    /// <summary>
    /// Interaction logic for LoadWindow.xaml
    /// </summary>
    public partial class LoadWindow : MahApps.Metro.Controls.MetroWindow
    {
        MainWindow mainWindow = new MainWindow();
        public LoadWindow()
        {
            InitializeComponent();
            this.Show();
            System.Threading.Thread.Sleep(5000);
            mainWindow.Show();
            this.Close();
        }
    }
}
