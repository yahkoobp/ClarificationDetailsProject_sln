// ----------------------------------------------------------------------------------------
// Project Name: ClarificationDetailsProject
// File Name: ViewModelBase.cs
// Description: // ViewModelBase class for implementing INotifyPropertyChanged functionality.
// Provides base property change notification mechanism for view models.
// Author: Yahkoob P
// Date: 27-10-2024
// ----------------------------------------------------------------------------------------
using System.ComponentModel;

namespace ClarificationDetailsProject.ViewModels
{
    /// <summary>
    /// A base view model class that implements the INotifyPropertyChanged interface.
    /// Used to provide property change notification for derived view models.
    /// </summary>
    public class ViewModelBase : INotifyPropertyChanged
    {
        /// <summary>
        /// Event triggered when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises the PropertyChanged event for a given property name.
        /// </summary>
        /// <param name="propertyName">Name of the property that changed.</param>
        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
