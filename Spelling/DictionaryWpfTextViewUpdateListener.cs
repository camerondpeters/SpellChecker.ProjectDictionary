using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using SpellChecker.Definitions;

namespace SpellChecker.ProjectDictionary
{
    
    [Export(typeof(IWpfTextViewCreationListener))]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    [ContentType("text")]
    sealed class DictionaryWpfTextViewUpdateListener : IWpfTextViewCreationListener
    {
        [Import] 
        IRegisterDictionaryTextBuffer _registerDictionaryTextBuffer = null;

        public void  TextViewCreated(IWpfTextView wpfTextView)
        {
            _registerDictionaryTextBuffer.RegisterTextBuffer(wpfTextView);
        }
    }
}
