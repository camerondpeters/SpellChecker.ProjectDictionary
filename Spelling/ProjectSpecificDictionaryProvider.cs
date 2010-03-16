using System;
using System.ComponentModel.Composition;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using SpellChecker.Definitions;

namespace SpellChecker.ProjectDictionary
{
    internal interface IRegisterDictionaryTextBuffer
    {
        void RegisterTextBuffer(IWpfTextView wpfTextView);
    }

    [Export(typeof(IBufferSpecificDictionaryProvider))]
    [Export(typeof(IRegisterDictionaryTextBuffer))]
    internal class ProjectSpecificDictionaryProvider : IBufferSpecificDictionaryProvider, IRegisterDictionaryTextBuffer
    {
        [Import]
        IVsEditorAdaptersFactoryService AdaptersFactory = null;

        [Import]
        ITextSearchService TextSearchService = null;

        // Links projects to individual dictionaries.
        private readonly Dictionary<String, IProjectSpecificDictionary> _dictionaries = new Dictionary<string, IProjectSpecificDictionary>(16);

        internal const string DictionaryFileName = "CustomDictionary.xml";

        private Events2 _events = null;
        private ProjectItemsEvents _ProjectItemsEvents = null;
        private SolutionEvents _SolutionEvents = null;

        internal ProjectSpecificDictionaryProvider()
        {
            var dte2 = (DTE2)Package.GetGlobalService(typeof(SDTE));
            
            if (dte2 != null)
            {
                _events = dte2.Events as Events2;
                if (_events != null)
                {
                    _ProjectItemsEvents = _events.ProjectItemsEvents;
                    _ProjectItemsEvents.ItemRenamed += ProjectItemsEvents_ItemRenamed;
                    _ProjectItemsEvents.ItemAdded += ProjectItemsEvents_ItemAddedOrRemoved;
                    _ProjectItemsEvents.ItemRemoved += ProjectItemsEvents_ItemAddedOrRemoved;

                    _SolutionEvents = _events.SolutionEvents;
                    _SolutionEvents.AfterClosing += SolutionEvents_AfterClosing;
                    _SolutionEvents.ProjectRenamed += SolutionEvents_ProjectRenamed;
                    _SolutionEvents.ProjectRemoved += SolutionEvents_ProjectRemoved;
                }
            }

        }

        void SolutionEvents_AfterClosing()
        {
            // Release all the dictionaries
            _dictionaries.Clear();
        }

        private void ProjectItemsEvents_ItemAddedOrRemoved(ProjectItem projectitem)
        {
            if (String.Compare(projectitem.Name, DictionaryFileName, StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                var dictionary = GetExistingSpellingDictionaryService(projectitem);
                if (dictionary != null)
                    dictionary.Reload(resetCreateFlag: true);
            }
        }

        private void ProjectItemsEvents_ItemRenamed(ProjectItem projectitem, string oldname)
        {
            // Is the before or after file name CustomDictionary.XML?
            if ((String.Compare(projectitem.Name, DictionaryFileName, StringComparison.InvariantCultureIgnoreCase) != 0) &&
                (String.Compare(oldname, DictionaryFileName, StringComparison.InvariantCultureIgnoreCase) != 0)) return;

            var dictionary = GetExistingSpellingDictionaryService(projectitem);
            if (dictionary != null)
            {
                // Find the view the changed project, if there is one.
                IWpfTextView wpfTextView = null;
                var iVsTextView = WpfTextViewHelpers.GetIVsTextView(dictionary.DictionaryFilePath);
                if (iVsTextView != null)
                    wpfTextView = AdaptersFactory.GetWpfTextView(iVsTextView);

                if (wpfTextView != null)
                    dictionary.ConnectBufferListener(wpfTextView.TextBuffer);

                dictionary.Reload(resetCreateFlag: true);
            }
        }

        private void SolutionEvents_ProjectRemoved(Project project)
        {
            if (String.IsNullOrEmpty(project.FullName)) return;

            // Look up the dictionary service in the hash table, and remove the reference.
            if (_dictionaries.Remove(project.FullName))
                Debug.WriteLine("Removed DictionaryService for: {0}", project.FullName);
        }

        private void SolutionEvents_ProjectRenamed(Project project, string oldname)
        {
            if (String.IsNullOrEmpty(oldname)) return;

            IProjectSpecificDictionary projectSpecificDictionary;
            if (_dictionaries.TryGetValue(oldname, out projectSpecificDictionary))
            {
                // Remove the reference off the old name
                _dictionaries.Remove(oldname);

                // Add the reference to the new name
                if (!String.IsNullOrEmpty(project.FullName)) 
                    _dictionaries.Add(project.FullName, projectSpecificDictionary);
            }
        }

        /// <summary>
        /// Called whenever a new view is created.  
        /// This method checks to see if it is a custom dictionary, and if it is, it gets attached.
        /// </summary>
        /// <param name="wpfTextView"></param>
        public void RegisterTextBuffer(IWpfTextView wpfTextView)
        {
            // If this isn't the dictionary, we don't care.
            var filePath = wpfTextView.GetFilePath();
            var fileName = Path.GetFileName(filePath);
            if (String.Compare(fileName, DictionaryFileName, StringComparison.InvariantCultureIgnoreCase) != 0)
                return;

            // Get the dictionary;
            var spellingDictionaryService = GetSpellingDictionaryService(fileName);

            // Do we have a project dictionary?
            var projectSpellingDictionaryService = spellingDictionaryService as IProjectSpecificDictionary;

            // Assign the buffer
            if (projectSpellingDictionaryService != null)
                projectSpellingDictionaryService.ConnectBufferListener(wpfTextView.TextBuffer);
        }

        /// <summary>
        /// Get's the IProjectSpecificDictionary for the project, if it exists.
        /// Return's null otherwise.
        /// </summary>
        private IProjectSpecificDictionary GetExistingSpellingDictionaryService(ProjectItem projectItem)
        {
            if ((projectItem != null) && (projectItem.ContainingProject != null))
            {
                IProjectSpecificDictionary spellingDictionary;
                if (_dictionaries.TryGetValue(projectItem.ContainingProject.FullName, out spellingDictionary))
                    return spellingDictionary;
            }

            return null;
        }

        /// <summary>
        /// Find a dictionary service for this textBuffer. Must always return something.
        /// </summary>
        /// <param name="textBuffer"></param>
        /// <returns></returns>
        public ISpellingDictionary GetDictionary(ITextBuffer textBuffer)
        {
            var fileName = textBuffer.GetFileName();
            return GetSpellingDictionaryService(fileName);
        }

        private ISpellingDictionary GetSpellingDictionaryService(string fileName)
        {
            var containingProject = WpfTextViewHelpers.GetContainingProject(fileName);

            Debug.WriteLine("Providing dictionary for: <{0}> in project <{1}>", fileName, containingProject != null && !String.IsNullOrEmpty(containingProject.FullName) ? containingProject.FullName : "MiscFiles");

            // If we are part of a collection of miscellaneous files, just return the standard spelling service.
            if ((containingProject == null) || (String.IsNullOrEmpty(containingProject.FullName)))
            {
                Debug.WriteLine("Misc files... returning null.");
                return null;
            }

            // We are tracking the dictionary services by the name of the project.

            // Does a dictionary exist, already?
            IProjectSpecificDictionary spellingDictionary;
            if (!_dictionaries.TryGetValue(containingProject.FullName, out spellingDictionary))
            {
                var projectDirectory = Path.GetDirectoryName(containingProject.FullName);
                var dictionaryName = Path.Combine(projectDirectory, DictionaryFileName);

                // No project level dictionary, so we need to create one.
                spellingDictionary = new ProjectSpecificDictionary(containingProject, AdaptersFactory,
                                                                                 TextSearchService, dictionaryName);
                Debug.WriteLine(String.Format("Created dictionary for <{0}>", containingProject.FullName));
                _dictionaries.Add(containingProject.FullName, spellingDictionary);
            }

            return spellingDictionary as ISpellingDictionary;
        }
    }
}

