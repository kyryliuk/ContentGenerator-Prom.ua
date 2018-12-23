using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    /// Interaction logic for TestRegexWindow.xaml
    /// </summary>
    public partial class TestRegexWindow : Window
    {
        public TestRegexWindow()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            OutputResult.Clear();
            string s = InputText.Text;
            Regex regex = new Regex(@"" + InputRegex.Text + "");
            MatchCollection matches = regex.Matches(s);
            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                    OutputResult.Text += (match.Value) + "\n";
            }
            else
            {
                MessageBox.Show("Совпадений не найдено");
            }
        }
    }
}
