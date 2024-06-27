using System.Collections.Generic;

namespace WarningAndErrorService
{
    /// <summary>
    /// Provides a centralized service for logging informational messages, warnings, and errors.  
    /// </summary>
    public class WarningService
    {
        /// <summary>
        /// Stores informational messages.
        /// </summary>
        private List<string> _infoMessages = new List<string>();

        /// <summary>
        /// Stores warning messages.
        /// </summary>
        private List<string> _warnings = new List<string>();

        /// <summary>
        /// Stores error messages.
        /// </summary>
        private List<string> _errors = new List<string>();

        /// <summary>
        /// Adds an informational message to the service.
        /// </summary>
        /// <param name="message">The informational message to add.</param>
        public void AddInfo(string message)
        {
            _infoMessages.Add(message);
        }

        /// <summary>
        /// Adds a warning message to the service.
        /// </summary>
        /// <param name="message">The warning message to add.</param>
        public void AddWarning(string message)
        {
            _warnings.Add(message);
        }

        /// <summary>
        /// Adds an error message to the service.
        /// </summary>
        /// <param name="message">The error message to add.</param>
        public void AddError(string message)
        {
            _errors.Add(message);
        }

        /// <summary>
        /// Indicates whether any critical errors have been logged.
        /// </summary>
        /// <returns>True if no errors are present, otherwise false.</returns>
        public bool IsErrorFree()
        {
            return _errors.Count == 0;
        }

        /// <summary>
        /// Indicates whether any errors have been logged.
        /// </summary>
        /// <returns>True if errors are present, otherwise false.</returns>
        public bool HasErrors()
        {
            return _errors.Count > 0;
        }

        /// <summary>
        /// Retrieves a list of informational messages.
        /// </summary>
        /// <returns>A copy of the list of informational messages.</returns>
        public List<string> GetInfos()
        {
            return new List<string>(_infoMessages);
        }

        /// <summary>
        /// Retrieves a list of warning messages.
        /// </summary>
        /// <returns>A copy of the list of warning messages.</returns>
        public List<string> GetWarnings()
        {
            return new List<string>(_warnings);
        }

        /// <summary>
        /// Retrieves a list of error messages.
        /// </summary>
        /// <returns>A copy of the list of error messages.</returns>
        public List<string> GetErrors()
        {
            return new List<string>(_errors);
        }
    }
}