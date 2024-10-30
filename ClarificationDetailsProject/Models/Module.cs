// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: Module.cs
// Description: Defines a class for Module
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------
using System.ComponentModel;

namespace ClarificationDetailsProject.Models
{
    /// <summary>
    /// Represents a module with a name and a checked status, typically used for filtering or categorization purposes.
    /// Implements INotifyPropertyChanged to notify the UI of changes to its properties.
    /// </summary>
    public class Module : INotifyPropertyChanged
    {
        /// <summary>
        /// Gets or sets the name of the module.
        /// </summary>
        public string Name { get; set; }

        private bool _isChecked;

        /// <summary>
        /// Gets or sets a value indicating whether the module is selected or checked.
        /// Notifies the UI of any changes.
        /// </summary>
        public bool IsChecked 
        {
            get
            {
                return _isChecked;
            }
            set
            {
                _isChecked = value;
                OnPropertyChanged(nameof(IsChecked));
            }
        }

        /// <summary>
        /// Occurs when a property value changes, used to notify the UI of property changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises the PropertyChanged event for the specified property.
        /// </summary>
        /// <param name="propertyName">The name of the property that changed.</param>
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
