// ----------------------------------------------------------------------------------------
// Project Name: PPTMerger
// File Name: MainWindow.cs
// Description: Code behind for MainWindow.xaml
// Author: Yahkoob P
// Date: 11-12-2024
// ----------------------------------------------------------------------------------------
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace PPTMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private object _draggedItem = new object();
        public MainWindow()
        {
            InitializeComponent();
            //set the data context for main window
            this.DataContext = new MainViewModel();
        }
        /// <summary>
        /// To handle previewMouseLeftButtonDown event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;

            // Walk up the visual tree to check if the source is within a Button
            while (source != null && !(source is ListBoxItem))
            {
                if (source is Button)
                {
                    // Allow the button to handle the event
                    return;
                }

                source = VisualTreeHelper.GetParent(source);
            }
            var listBox = sender as ListBox;
            if (listBox != null)
            {
                // Get the item under the mouse pointer
                var position = e.GetPosition(listBox);
                _draggedItem = GetListBoxItemUnderMouse(listBox, position) as string;  // assuming string collection

                if (_draggedItem != null)
                {
                    // Start the drag operation (this will call DragDrop.DoDragDrop)
                    DragDrop.DoDragDrop(listBox, _draggedItem, DragDropEffects.Move);
                }
            }
        }

        /// <summary>
        /// To handle ListBox drop event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var source = e.OriginalSource as DependencyObject;

            // Walk up the visual tree to check if the source is within a Button
            while (source != null && !(source is ListBoxItem))
            {
                if (source is Button)
                {
                    // Allow the button to handle the event
                    return;
                }

                source = VisualTreeHelper.GetParent(source);
            }
           // e.Handled = false;
            if (_draggedItem != null)
            {
                var listBox = sender as ListBox;
                var items = listBox.ItemsSource as ObservableCollection<string>;

                var targetItem = GetListBoxItemUnderMouse(listBox, e.GetPosition(listBox)) as string;

                if (targetItem != null && !ReferenceEquals(_draggedItem, targetItem))
                {
                    int oldIndex = items.IndexOf((string)_draggedItem);
                    int newIndex = items.IndexOf(targetItem);

                    if (oldIndex >= 0 && newIndex >= 0)
                    {
                        // Move the dragged item to the new position
                        items.Move(oldIndex, newIndex);
                    }
                }

                _draggedItem = null;
            }
        }

        /// <summary>
        /// method to get the listbox item under the mouse
        /// </summary>
        /// <param name="listBox"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        private object GetListBoxItemUnderMouse(ListBox listBox, Point position)
        { 

            var hitTestResult = VisualTreeHelper.HitTest(listBox, position);
            if (hitTestResult != null)
            {
                var dependencyObject = hitTestResult.VisualHit;
                while (dependencyObject != null && !(dependencyObject is ListBoxItem))
                {
                    dependencyObject = VisualTreeHelper.GetParent(dependencyObject);
                }

                if (dependencyObject is ListBoxItem listBoxItem)
                {
                    return listBox.ItemContainerGenerator.ItemFromContainer(listBoxItem);
                }
            }
            return null;
        }
    }
}
