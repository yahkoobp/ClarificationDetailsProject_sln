using System;
using System.Collections;
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
           DependencyProperty.Register("Items", typeof(IEnumerable), typeof(SummaryGridControl), new PropertyMetadata(null));
        public IEnumerable Items
        {
            get { return (IEnumerable)GetValue(ItemsProperty); }
            set { SetValue(ItemsProperty, value); }
        }

        public static readonly DependencyProperty StatusProperty =
           DependencyProperty.Register("Status", typeof(IEnumerable), typeof(SummaryGridControl), new PropertyMetadata(null));
        public IEnumerable Status
        {
            get { return (IEnumerable)GetValue(ItemsProperty); }
            set { SetValue(ItemsProperty, value); }
        }
    }
}
