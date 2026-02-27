using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SolidWorksTankDesign.Helpers
{
    public static class CompartmentsMappingHelper
    {
        // The static dictionary to hold the mappings
        public static Dictionary<Guid, (Compartment compartment, CompartmentConfiguration config)> _idMapping
            = new Dictionary<Guid, (Compartment, CompartmentConfiguration)>();

        public static Dictionary<Guid, CompartmentConfiguration> _originalConfigurations = new Dictionary<Guid, CompartmentConfiguration>();

        /// <summary>
        /// Adds a new compartment and configuration pair to the mapping.
        /// If the ID already exists, it updates the existing entry.
        /// </summary>
        /// <param name="compartment">The Compartment object</param>
        /// <param name="configuration">The CompartmentConfiguration object</param>
        public static void AddOrUpdateMapping(Compartment compartment, CompartmentConfiguration configuration)
        {
            if (compartment == null || configuration == null)
                throw new ArgumentNullException("Compartment and Configuration cannot be null.");

            if (compartment._compartmentSettings.ID != configuration.ID)
                throw new InvalidOperationException("Compartment ID and Configuration ID must match.");

            _idMapping[compartment._compartmentSettings.ID] = (compartment, configuration);
        }

        /// <summary>
        /// Retrieves the compartment and configuration pair by ID.
        /// </summary>
        /// <param name="id">The unique ID of the compartment</param>
        /// <returns>A tuple of Compartment and CompartmentConfiguration</returns>
        public static (Compartment compartment, CompartmentConfiguration config)? GetMappingById(Guid id)
        {
            if (_idMapping.TryGetValue(id, out var mapping))
                return mapping;

            return null; // Return null if the ID is not found
        }

        /// <summary>
        /// Returns the entire mapping dictionary.
        /// </summary>
        public static Dictionary<Guid, (Compartment compartment, CompartmentConfiguration config)> GetAllMappings()
        {
            return new Dictionary<Guid, (Compartment, CompartmentConfiguration)>(_idMapping);
        }

        /// <summary>
        /// Deletes a mapping by ID.
        /// </summary>
        /// <param name="id">The ID of the mapping to delete</param>
        public static void RemoveMapping(Guid id)
        {
            if (_idMapping.ContainsKey(id))
            {
                _idMapping.Remove(id);
            }
        }

        /// <summary>
        /// Clears all mappings.
        /// </summary>
        public static void ClearMappings()
        {
            _idMapping.Clear();
        }

        /// <summary>
        /// Saves a snapshot of the current configurations.
        /// </summary>
        public static void SaveOriginalConfigurations()
        {
            _originalConfigurations.Clear();

            Dictionary<Guid, (Compartment compartment, CompartmentConfiguration config)> mappings = CompartmentsMappingHelper.GetAllMappings();

            foreach (var mapping in mappings)
            {
                // Deep copy the configuration to avoid accidental reference updates
                _originalConfigurations[mapping.Key] = mapping.Value.config.DeepClone();
            }
        }

        public static List<Guid> GetModifiedConfigurations()
        {
            Dictionary<Guid, (Compartment compartment, CompartmentConfiguration config)> mappings = CompartmentsMappingHelper.GetAllMappings();

            List<Guid> modifiedIds = new List<Guid>();

            foreach (var mapping in mappings)
            {
                var currentConfig = mapping.Value.config;
                if (_originalConfigurations.TryGetValue(mapping.Key, out var originalConfig))
                {
                    // Compare the original and current configurations
                    if (!AreConfigurationsEqual(originalConfig, currentConfig))
                    {
                        modifiedIds.Add(mapping.Key);
                    }
                }
                else
                {
                    // If no original configuration exists, treat it as modified
                    modifiedIds.Add(mapping.Key);
                }
            }

            return modifiedIds;
        }

        /// <summary>
        /// Compares two configurations for equality.
        /// </summary>
        private static bool AreConfigurationsEqual(CompartmentConfiguration original, CompartmentConfiguration current)
        {
            if (original == null || current == null)
                return false;

            return original.Volume == current.Volume &&
                   original.Length == current.Length &&
                   original.InternalSurfaceTreatment == current.InternalSurfaceTreatment &&
                   original.Amount == current.Amount &&
                   original.AmountUnits == current.AmountUnits &&
                   original.LeftDishedEndAlignment == current.LeftDishedEndAlignment &&
                   original.LeftEndConnection == current.LeftEndConnection;
        }

        public static List<Guid> GetDeletedConfigurations()
        {
            List<Guid> deletedIds = new List<Guid>();

            foreach (var originalEntry in _originalConfigurations)
            {
                // Check if the current mapping no longer contains this ID
                if (!_idMapping.ContainsKey(originalEntry.Key))
                {
                    deletedIds.Add(originalEntry.Key);
                }
            }

            return deletedIds;
        }

        public static List<(CompartmentConfiguration config, int index)> GetAddedConfigurations()
        {
            List<(CompartmentConfiguration config, int index)> newConfigurations = new List<(CompartmentConfiguration config, int index)>();

            for (int i = 0; i <  SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Count; i++)
            {
                CompartmentConfiguration compartmentConfiguration = SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i];

                // Check if the ID from mapping is not present in original configurations
                if (!_originalConfigurations.ContainsKey(compartmentConfiguration.ID))
                {
                    newConfigurations.Add((compartmentConfiguration, i));
                }
            }

            return newConfigurations;
        }

        public static List<(CompartmentConfiguration config, int index)> GetSwitchedConfigurations(
            ObservableCollection<CompartmentConfiguration> updatedList)
        {
            List<(CompartmentConfiguration, int) > switchedCompartmentConfigs = new List<(CompartmentConfiguration, int)>();

            List<(Guid, CompartmentConfiguration)> originalConfigsList = new List<(Guid, CompartmentConfiguration)>();
            foreach (var originalConfig in _originalConfigurations)
            {
                (Guid, CompartmentConfiguration) configToAdd = ( originalConfig.Value.ID, originalConfig.Value );
                originalConfigsList.Add(configToAdd);
            }
               
            // Build a quick ID => newIndex map for the updated list
            Dictionary<Guid, int> newIndexMap = new Dictionary<Guid, int>();
            for (int i = 0; i < updatedList.Count; i++)
            {
                newIndexMap[updatedList[i].ID] = i;
            }

            // For each original config, find if it still exists in the updated list
            // and see if its position changed.
            for (int i = 0; i < originalConfigsList.Count; i++)
            {
                Guid originalId = originalConfigsList[i].Item1;
                CompartmentConfiguration originalConfig = originalConfigsList[i].Item2;

                if (newIndexMap.TryGetValue(originalId, out int newIndex))
                {
                    if (i != newIndex)
                        switchedCompartmentConfigs.Add((originalConfig, newIndex));
                }
            }

            return switchedCompartmentConfigs;
        }

        public static void DeleteCompartmentsByDeletedConfigurationsFromMapping(List<Guid> deletedCompartmentsIds)
        {
            foreach (var id in deletedCompartmentsIds)
            {
                // Safely remove the mapping
                if (_idMapping.ContainsKey(id))
                {
                    _idMapping.Remove(id);
                }
            }
        }
    }
}

