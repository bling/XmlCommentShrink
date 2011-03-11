using System;
using System.ComponentModel.Composition;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace XmlCommentShrink
{
    [Export(typeof(IWpfTextViewCreationListener))]
    [ContentType("code")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    public class Listener : IWpfTextViewCreationListener
    {
        [Import]
        private IClassificationFormatMapService _formatMapService;

        public void TextViewCreated(IWpfTextView textView)
        {
            textView.Properties.GetOrCreateSingletonProperty(
                () => new XmlCommentShrinker(textView, _formatMapService.GetClassificationFormatMap(textView)));
        }
    }

    internal class XmlCommentShrinker
    {
        private bool _updating;
        private readonly IClassificationFormatMap _formatMap;

        public XmlCommentShrinker(ITextView view, IClassificationFormatMap formatMap)
        {
            _formatMap = formatMap;
            _formatMap.ClassificationFormatMappingChanged += (sender, e) => ShrinkComments();

            view.GotAggregateFocus += OnGotAggregateFocus;
            ShrinkComments();
        }

        private void OnGotAggregateFocus(object sender, EventArgs e)
        {
            ((ITextView)sender).GotAggregateFocus -= OnGotAggregateFocus;
            ShrinkComments();
        }

        private void ShrinkComments()
        {
            if (_updating)
                return;

            try
            {
                _updating = true;

                foreach (var classification in from c in _formatMap.CurrentPriorityOrder
                                               where c != null
                                               let name = c.Classification.ToLowerInvariant()
                                               where name.Contains("doc")
                                               select c)
                {
                    Shrink(classification);
                }
            }
            finally
            {
                _updating = false;
            }
        }

        private void Shrink(IClassificationType classification)
        {
            var properties = _formatMap.GetTextProperties(classification);
            var typeface = properties.Typeface;

            if (typeface.Style == FontStyles.Italic)
                return;

            _formatMap.SetTextProperties(
                classification,
                properties
                    .SetTypeface(new Typeface(typeface.FontFamily, FontStyles.Italic, typeface.Weight, typeface.Stretch))
                    .SetFontRenderingEmSize(6));
        }
    }
}