using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace BCK
{
    public class ViewModelBase : INotifyPropertyChanged, INotifyDataErrorInfo
    {
        #region INotifyPropertyChanged interface
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        #region INotifyDataErrorInfo interface
        public event EventHandler<DataErrorsChangedEventArgs> ErrorsChanged;

        private readonly Dictionary<string, string> ValidationErrors = new Dictionary<string, string>();

        public IEnumerable GetErrors(string propertyName)
        {
            return ValidationErrors.ContainsKey(propertyName)
                ? ValidationErrors[propertyName]
                : null;
        }

        public bool HasErrors
        {
            get
            {
                return ValidationErrors.Count > 0;
            }
        }

        public bool IsValid
        {
            get { return !HasErrors; }
        }
        #endregion

        #region ViewModelBase
        protected void HandleValidationError(bool isValid, string propertyName, string validationErrorDescription)
        {
            if (!string.IsNullOrEmpty(propertyName))
            {
                if (isValid)
                {
                    if (ValidationErrors.ContainsKey(propertyName))
                    {
                        ValidationErrors.Remove(propertyName);
                    }
                }
                else
                {
                    if (!ValidationErrors.ContainsKey(propertyName))
                    {
                        ValidationErrors.Add(propertyName, validationErrorDescription);
                    }

                    else
                    {
                        ValidationErrors[propertyName] = validationErrorDescription;
                    }
                        
                    throw new ArgumentException(validationErrorDescription);
                }
            }
        }

        protected bool HasPropertyError(string propertyName)
        {
            return ValidationErrors.ContainsKey(propertyName);
        }
        #endregion
    }
}
