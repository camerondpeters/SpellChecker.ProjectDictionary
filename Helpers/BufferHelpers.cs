using System;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;

namespace SpellChecker.ProjectDictionary
{
    internal static class BufferHelpers
    {
        public static string GetFileName(this ITextBuffer textBuffer)
        {
            IVsTextBuffer bufferAdapter;
            textBuffer.Properties.TryGetProperty(typeof(IVsTextBuffer), out bufferAdapter);

            if (bufferAdapter != null)
            {
                var persistFileFormat = bufferAdapter as IPersistFileFormat;
                if (persistFileFormat != null)
                {
                    string ppzsFilename;
                    uint iii;
                    persistFileFormat.GetCurFile(out ppzsFilename, out iii);
                    return ppzsFilename;
                }
            }
            return String.Empty;
        }
    }
}