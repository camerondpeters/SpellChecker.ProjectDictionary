using Microsoft.VisualStudio.Text;

namespace SpellChecker.ProjectDictionary
{
    internal interface IProjectSpecificDictionary
    {
        void ConnectBufferListener(ITextBuffer dictionaryTextBuffer);
        string DictionaryFilePath { get; }
        void Reload(bool resetCreateFlag);
    }
}