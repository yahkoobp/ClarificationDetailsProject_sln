using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using ClarificationDetailsProject.Models;

namespace ClarificationDetailsProject.CustomControls
{
    /// <summary>
    /// Interaction logic for SummaryGridControl.xaml
    /// </summary>
    public partial class SummaryGridControl : UserControl
    {
        public SummaryGridControl()
        {
            InitializeComponent();
        }
        public static readonly DependencyProperty ItemsProperty =
           DependencyProperty.Register("Items", typeof(ObservableCollection<Summary>), typeof(SummaryGridControl), new PropertyMetadata(null));
        public ObservableCollection<Summary> Items
        {
            get { return (ObservableCollection<Summary>)GetValue(ItemsProperty); }
            set { SetValue(ItemsProperty, value); }
        }
    }
}
