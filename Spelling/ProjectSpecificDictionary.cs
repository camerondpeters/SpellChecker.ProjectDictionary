using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Threading;
using System.Xml;
using System.Text;
using System.Xml.Linq;
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
    internal class ProjectSpecificDictionary : IProjectSpecificDictionary, ISpellingDictionary
    {
        private Project _project;
        private readonly string _dictionaryFilePath;
        private readonly IVsEditorAdaptersFactoryService _adaptersFactory;
        private readonly ITextSearchService _textSearchService;

        private FileSystemWatcher _dictionaryFileWatcher;
        private DispatcherTimer _deferTimer;
        private const int DeferredReloadDelay = 2500;  // Time to wait after the file changes on disk, before reloading it.

        private SortedSet<string> _projectDictionaryWords = new SortedSet<string>();
        private SortedSet<string> _ignoreWords = new SortedSet<string>();

        internal const string EmptyDictionary = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
                                       "<Dictionary>\r\n" +
                                       "\t<Words>\r\n" +
                                       "\t\t<Recognized>\r\n" +
                                       /* "\t\t\t<Word>Normalise</Word>\n" + // Add one British word. */
                                       "\t\t</Recognized>\r\n" +
                                       "\t</Words>\r\n" +
                                       "</Dictionary>\r\n";

        internal  ProjectSpecificDictionary(Project project, IVsEditorAdaptersFactoryService adaptersFactory, ITextSearchService textSearchService, string dictionaryFileName)
        {
            if (String.IsNullOrEmpty(dictionaryFileName))
                throw new ArgumentException("dictionaryFileName");

            _project = project;
            _adaptersFactory = adaptersFactory;
            _textSearchService = textSearchService;

            // The dictionary file name
            _dictionaryFilePath = dictionaryFileName;

            _dictionaryFileWatcher = new FileSystemWatcher(Path.GetDirectoryName(_dictionaryFilePath), Path.GetFileName(_dictionaryFilePath))
                                         {
                                             IncludeSubdirectories = false,
                                             EnableRaisingEvents = true,
                                         };
            _dictionaryFileWatcher.Changed += DictionaryFileWatcherChangedCreatedOrDeleted;
            _dictionaryFileWatcher.Created += DictionaryFileWatcherChangedCreatedOrDeleted;
            _dictionaryFileWatcher.Deleted += DictionaryFileWatcherChangedCreatedOrDeleted;
            _dictionaryFileWatcher.Renamed += DictionaryFileWatcherRenamed;

            _deferTimer = new DispatcherTimer(DispatcherPriority.ApplicationIdle)
                              {
                                  Interval = new TimeSpan(DeferredReloadDelay)
                              };
            _deferTimer.Tick += (s, e) =>
                                    {
                                        _deferTimer.Stop();
                                        Reload(false);
                                    };

            // Load the dictionary... initial load should never attempt to get the ITextBuffer, because
            // it will recurse forever.
            var dictionaryProject = WpfTextViewHelpers.GetContainingProject(DictionaryFilePath);
            if (project == dictionaryProject)
                LoadXmlDictionary(DictionaryFilePath);
        }

        void DictionaryFileWatcherRenamed(object sender, RenamedEventArgs e)
        {
            DeferedReload();
        }

        void DictionaryFileWatcherChangedCreatedOrDeleted(object sender, FileSystemEventArgs e)
        {
            DeferedReload();
        }

        // SystemFileWatcher events can come in clusters, so we want to make sure
        // that they've all happened before we reload the dictionary.
        void DeferedReload()
        {
            lock (_deferTimer)
                _deferTimer.Start();
        }

        private void ClearXmlDictionary()
        {
            if (_projectDictionaryWords.Count == 0) return;

            lock (_projectDictionaryWords)
                _projectDictionaryWords.Clear();

            // Recheck everything.
            RaiseSpellingChangedEvent(null);
        }

        public string DictionaryFilePath
        {
            get { return _dictionaryFilePath; }
        }

        // This method should never allow any exceptions to escape.
        private void LoadXmlDictionary(string fileName)
        {
            try
            {
                LoadXmlDictionary(XDocument.Load(fileName));
            }
            catch (Exception)
            {}
        }

        // This method should never allow any exceptions to escape.
        private void LoadXmlDictionary(ITextSnapshot snapshot)
        {
            try
            {
                LoadXmlDictionary(XDocument.Parse(snapshot.GetText()));
            }
            catch (Exception)
            {}
        }

        // This method should never allow any exceptions to escape.
        private void LoadXmlDictionary(XDocument dictionary)
        {
            try
            {
                // I am sure that a more clever linq expression is possible here.
                IEnumerable<string> words = null;
                var recognizedElement = dictionary.Descendants("Recognized").FirstOrDefault();
                if (recognizedElement != null)
                    words = from w in recognizedElement.Descendants("Word") where !String.IsNullOrEmpty(w.Value) select w.Value;

                // We have no valid dictionary, so don't change anything.
                if (words == null) return;

                var newWords = new SortedSet<string>(words);

                // Get a copy of the ignore words
                var ignorewords = _projectDictionaryWords;

                // Replace the ignorewords with our new words.  Not sure if the lock is necessary??
                lock (_projectDictionaryWords)
                    _projectDictionaryWords = newWords;

                // If we get to here, we have successfully loaded a dictionary.

                // Now figure out the added words.
                var exceptionWords = new SortedSet<string>(newWords);
                exceptionWords.ExceptWith(ignorewords);

                // Added words.
                foreach (var w in exceptionWords)
                    RaiseSpellingChangedEvent(w);

                // Figure out the deleted words.
                ignorewords.ExceptWith(newWords);

                // If any words are deleted, all text must be dirtied and rechecked
                if (ignorewords.Count > 0)
                    RaiseSpellingChangedEvent(null);
            }
            catch (Exception)
            {
                // Do nothing, do not update the ignore words.
            }
        }

        #region Implementation of ISpellingDictionary

        /// <summary>
        /// Add the given word to the dictionary, so that it will no longer show up as an
        /// incorrect spelling.
        /// </summary>
        /// <param name="word">The word to add to the dictionary.</param>
        public bool AddWordToDictionary(string word)
        {
            // Nothing to add...
            if (string.IsNullOrEmpty(word))
                return true;

            // Is there dictionary file in a project, if not, prompt to create one (once)
            if ((WpfTextViewHelpers.GetContainingProject(DictionaryFilePath) == null) && !PromptToCreateLocalDictionary())
                return false;

            // Check if it's in the dictionary...
            if (!_projectDictionaryWords.Contains(word))
            {
                lock (_projectDictionaryWords)
                    _projectDictionaryWords.Add(word);

                var added = AddWordToCustomDictionaryXML(word);
                if (added)
                    RaiseSpellingChangedEvent(word);

                return added;
            }

            // Word was already in the dictionary...
            return true;
        }

        private bool _promptedToCreateLocalDictionary = false;
        private bool PromptToCreateLocalDictionary()
        {
            if (_promptedToCreateLocalDictionary == true)
                return false;

            // We are just going to do this once.
            _promptedToCreateLocalDictionary = true;

            // Should we create a dictionary??
            var iVsUIShell = (IVsUIShell)Package.GetGlobalService(typeof(SVsUIShell));
            if (VsShellUtilities.PromptYesNo("Add a CustomDictionary.XML file to this project?", "Spell Checker", OLEMSGICON.OLEMSGICON_QUERY, iVsUIShell))
            {
                // If the file exists, prompt to use it
                if (!File.Exists(DictionaryFilePath) || (!VsShellUtilities.PromptYesNo("CustomDictionary.XML file already exists. Use existing file?\r\n(Select <No> to erase the current file and create a new one)", "Spell Checker", OLEMSGICON.OLEMSGICON_QUERY, iVsUIShell)))
                {
                    // Create the dictionary file.
                    try
                    {
                        using (var streamWriter = new StreamWriter(DictionaryFilePath, false))
                            streamWriter.Write(EmptyDictionary);
                    }
                    catch (Exception e)
                    {
                        // Show a message box...
                        Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)Package.GetGlobalService(typeof(SDTE)); ;
                        Microsoft.VisualStudio.Shell.ServiceProvider serviceProvider = new Microsoft.VisualStudio.Shell.ServiceProvider(sp);

                        VsShellUtilities.ShowMessageBox(serviceProvider, e.Message, "Can't create CustomDictionary.XML",
                                                        OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                        // Failed...)
                        return false;   
                    }
                }

                // Add the dictionary to the project.
                _project.ProjectItems.AddFromFile(DictionaryFilePath);

                // Success!!
                return true;
            }

            return false;
        }

        private bool AddWordToCustomDictionaryXML(string word)
        {
            // Check out the dictionary file.
            var dte2 = (DTE2)Package.GetGlobalService(typeof(SDTE));
            var sourceControl = dte2.SourceControl;

            // Check out the dictionary file, if required.
            if ((sourceControl != null) && (sourceControl.IsItemUnderSCC(DictionaryFilePath)))
            {
                if (!sourceControl.IsItemCheckedOut(DictionaryFilePath))
                    sourceControl.CheckOutItem(DictionaryFilePath);
            }

            var textView = WpfTextViewHelpers.GetIVsTextView(DictionaryFilePath);
            if (textView != null)
            {
                var wpfTextView = _adaptersFactory.GetWpfTextView(textView);

                // Do we have the file in memory?))
                if ((wpfTextView != null) && (wpfTextView.TextBuffer != null))
                {
                    var textBuffer = wpfTextView.TextBuffer;

                    // Find </Recognized>
                    var snapshot = _textSearchService.FindNext(0, false,
                                                               new FindData("</Recognized>", textBuffer.CurrentSnapshot));
                    if (snapshot.HasValue)
                    {
                        // Move back until we find the beginning of the line, or the first non-whitespace character
                        var insertPos = snapshot.Value.Start;
                        var line = insertPos.GetContainingLine();

                        while ((insertPos > line.Start) && Char.IsWhiteSpace((insertPos - 1).GetChar()))
                            insertPos = insertPos - 1;

                        textBuffer.Insert(insertPos, String.Format("\t\t\t<Word>{0}</Word>\r\n", word));
                    }

                    return true;
                }
            }

            // Add the word to the dictionary, if required.
            try
            {
                var dictionary = XDocument.Load(DictionaryFilePath);
                var recognizedElement = dictionary.Descendants("Recognized").FirstOrDefault();
                if (recognizedElement != null)
                {
                    if ((from w in recognizedElement.Descendants("Word") where String.Compare(w.Value, word, false) == 0 select w).Count() == 0)
                        recognizedElement.Add(new XElement("Word", word));
                }
                dictionary.Save(DictionaryFilePath);

                return true;
            }
            catch (Exception)
            {
            }

            // If we get to here, we have failed
            return false;
        }

        /// <summary>
        /// Ignore the given word, but don't add it to the dictionary.
        /// </summary>
        /// <param name="word">The word to be ignored.</param>
        public bool IgnoreWord(string word)
        {
            if (!string.IsNullOrEmpty(word) && !_projectDictionaryWords.Contains(word))
            {
                lock (_ignoreWords)
                    _ignoreWords.Add(word);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Check the ignore dictionary for the given word.
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        public bool ShouldIgnoreWord(string word)
        {
            lock (_projectDictionaryWords)
                if (_projectDictionaryWords.Contains(word))
                    return true;

            lock (_ignoreWords)
                return _ignoreWords.Contains(word);
        }

        public event EventHandler<SpellingEventArgs> DictionaryUpdated;

        #endregion

        #region Implementation of IProjectSpecificDictionary

        public void ConnectBufferListener(ITextBuffer dictionaryTextBuffer)
        {
            if (dictionaryTextBuffer == null)
                throw new ArgumentNullException("dictionaryTextBuffer");

            // hook up new event handlers.
            BufferIdleEventUtil.AddBufferIdleEventListener(dictionaryTextBuffer, ReparseDictionary);
        }

        public void Reload(bool resetCreateLocalDictionaryFlag)
        {
            Debug.WriteLine("Reloading dictionary.");
            if (resetCreateLocalDictionaryFlag)
                _promptedToCreateLocalDictionary = false;

            // Look for the dictionary file in the project
            var project = WpfTextViewHelpers.GetContainingProject(DictionaryFilePath);
            if (project == null)
            {
                ClearXmlDictionary();
                return;
            }

            // Is the dictionary currently in a buffer?
            var iVsTextView = WpfTextViewHelpers.GetIVsTextView(DictionaryFilePath);
            if (iVsTextView != null)
            {
                var wpfTextView = _adaptersFactory.GetWpfTextView(iVsTextView);
                LoadXmlDictionary(wpfTextView.TextBuffer.CurrentSnapshot);
                Debug.WriteLine("Dictionary snapshot version: <{0}>", wpfTextView.TextBuffer.CurrentSnapshot.Version.VersionNumber);
                return;
            }

            // Load it from disk.
            LoadXmlDictionary(DictionaryFilePath);
        }

        /// <summary>
        /// This happens when the dictionary is loaded in memory, and it changes.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ReparseDictionary(object sender, EventArgs e)
        {
            var textBuffer = sender as ITextBuffer;

            // This seems highly unlikely, but we don't want our extension to crash...
            if (textBuffer == null) return;

            // Make sure this buffer is still the right one for our dictionary
            // (this can happen if a dictionary file is renamed to something else)
            if (String.Compare(textBuffer.GetFileName(), DictionaryFilePath, StringComparison.InvariantCultureIgnoreCase) == 0)
                LoadXmlDictionary(textBuffer.CurrentSnapshot);
        }

        #endregion

        #region Helpers

        void RaiseSpellingChangedEvent(string word)
        {
            var temp = DictionaryUpdated;
            if (temp != null)
                DictionaryUpdated(this, new SpellingEventArgs(word));
        }

        #endregion

        #region IDisposable
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _disposed;

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Unhook our file watcher
                    _dictionaryFileWatcher.Dispose();
                    _dictionaryFileWatcher = null;

                    // And the timer
                    _deferTimer = null;
                }

                _disposed = true;
            }
        }

        #endregion

    }
}
