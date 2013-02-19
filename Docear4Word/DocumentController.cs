using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

using Docear4Word.BibTex;

using Office;

using Word;

namespace Docear4Word
{
	public partial class DocumentController
	{
		const string DocearMarker = "Docear";
		const string CslCitationMarker = "CSL_CITATION";
		const string CslBibliographyMarker = "CSL_BIBLIOGRAPHY";
		const string AddInMarker = "ADDIN";
		const string InsertingBibliographyMessage = @"""Creating Bibliography...""";
		const string InsertingCitationMessage = @"""Creating Citation...""";
		const string DoNotModifyWarningMessage = "THIS CITATION DATA SHOULD NOT BE MANUALLY MODIFIED!!!";
		const string JSONWhitespace = " ";
		const string FieldCodeSeparator = "\v";

		const string StyleDocumentPropertyName = "Docear4Word_StyleTitle";
		const string DatabaseFilenameDocumentPropertyName = "Docear4Word_DatabaseFilename";

		static readonly Font FixedWidthCodeFont;

		static DocumentController()
		{
			FixedWidthCodeFont = new Font
			            	{
			            		Name = "Courier New",
			            		Size = 7.0f,
			            	};
		}

		readonly MainController mainController;
		readonly Document document;
		readonly DocumentProperties documentProperties;
		readonly Dictionary<string, JSInlineCitation> inlineCitationCache = new Dictionary<string, JSInlineCitation>();

		StyleInfo style;
		CiteProcRunner citeProc;
		bool isUpdating;
		StyleInfo duffStyleInfo;

		public DocumentController(MainController mainController, Document document)
		{
			this.mainController = mainController;
			this.document = document;

			documentProperties = (DocumentProperties) document.CustomDocumentProperties;

			var documentStyleTitle = GetDocumentStyleTitle();
			style = mainController.FindStyleByTitle(documentStyleTitle) ?? mainController.DefaultStyle;
		}

		string GetDocumentDatabaseFilename()
		{
			return DocumentHelper.GetCustomStringProperty(documentProperties, DatabaseFilenameDocumentPropertyName);
		}

		public void SetDocumentDatabaseFilename(string databaseFilename)
		{
			DocumentHelper.SetCustomProperty(documentProperties, DatabaseFilenameDocumentPropertyName, CustomPropertyType.String, databaseFilename);
		}

		string GetDocumentStyleTitle()
		{
			return DocumentHelper.GetCustomStringProperty(documentProperties, StyleDocumentPropertyName);
		}

		void SetDocumentStyleTitle(string styleTitle)
		{
			DocumentHelper.SetCustomProperty(documentProperties, StyleDocumentPropertyName, CustomPropertyType.String, styleTitle);
		}

		public bool IsUpdating
		{
			get { return isUpdating; }
		}

		public Document Document
		{
			get { return document; }
		}

		public BibTexDatabase GetDatabase()
		{
			if (Settings.Instance.AllowPerDocumentDatabases)
			{
				try
				{
					var documentDatabaseFilename = GetDocumentDatabaseFilename();

					if (documentDatabaseFilename != null)
					{
						var result = BibTexHelper.LoadBibTexDatabase(documentDatabaseFilename);

						if (result != null) return result;
					}
				}
				catch
				{}
			}

			return Settings.Instance.GetDefaultDatabase();
		}

		public void Activate()
		{
			// Stuff to get valid title here
		}

		public void UpdateDocumentProperties()
		{
			if (style == null || GetDocumentStyleTitle() != style.Title)
			{
				SetDocumentStyleTitle(style == null ? mainController.DefaultStyle.Title : style.Title);
			}
		}

		public StyleInfo Style
		{
			get { return style; }
			set
			{
				if (value == style) return;

				style = value;

				Refresh(false);
			}
		}

		internal CiteProcRunner CiteProc
		{
			get
			{
				if (citeProc == null)
				{
					var styleToUse = style ?? mainController.DefaultStyle;

					// Duff is duff - no use retrying
					if (styleToUse == duffStyleInfo)
					{
						return null;
					}

					try
					{
						citeProc = new CiteProcRunner(styleToUse, GetDatabase);
						duffStyleInfo = null;
					}
					catch
					{
						citeProc = null;
						duffStyleInfo = styleToUse;
					}
				}

				return citeProc;
			}
		}

		public void InsertCitationSequence(IEnumerable<EntryAndPagePair> itemSources, bool isLineSequence)
		{
			try
			{
				isUpdating = true;

				foreach (var itemSource in itemSources)
				{
					if (isLineSequence)
					{
						document.Application.Selection.TypeText("\v");
					}
					else
					{
						document.Application.Selection.TypeParagraph();
					}

					var citation = CreateInlineCitation(new[] { itemSource });

					InsertCitationCore(citation);
				}
			}
			finally
			{
				isUpdating = false;
			}
		}

		public void InsertCitation(List<EntryAndPagePair> entryAndPagePairs)
		{
			InsertCitationCore(CreateInlineCitation(entryAndPagePairs));
		}

		void InsertCitationCore(JSInlineCitation citation)
		{
			try
			{
				isUpdating = true;

				var selection = document.Application.Selection;
				selection.Collapse(WdCollapseDirection.wdCollapseEnd);
				var range = selection.Range;

				var inserter = new CitationInserter(this);
				inserter.InsertCitation(range, citation);
			}
			finally
			{
				isUpdating = false;
			}
		}

		public void EditCitation(Field field, List<EntryAndPagePair> entryAndPagePairs)
		{
			try
			{
				isUpdating = true;

				var inserter = new CitationInserter(this);

				inserter.EditCitation(field, CreateInlineCitation(entryAndPagePairs));
			}
			finally
			{
				isUpdating = false;
			}
		}

		public void InsertBibliography()
		{
			try
			{
				isUpdating = true;

				var selection = document.Application.Selection;
				selection.Collapse(WdCollapseDirection.wdCollapseEnd);
				var range = selection.Range;

				// Gotta clear the cache otherwise might show biblio
				Refresh(false);

				var inserter = new CitationInserter(this);
				inserter.InsertBibliography(range);
			}
			finally
			{
				isUpdating = false;
			}
		}

		JSInlineCitation CreateInlineCitation(IEnumerable<EntryAndPagePair> itemSources, object idToUse = null)
		{
			// ****IMPORTANT****
			// This is called from InsertCitationSequence, InsertCitation, EditCitation and CitationInserter.UpdateCitationsFromDatabase
			//
			// It is imperative that calls from the first 3 work on an empty CiteProc otherwise the cache gets used to create
			// the citation items. Other than first-use, this means using the item after CiteProc has seen it and maybe modified it
			// (it appears to change the Date Parts to strings in some cases)
			// The next refresh is then comparing incorrect JSON and will want to update it from the database.
			//
			// (CitationInserter.UpdateCitationsFromDatabase calls here but this is always within a Refresh which means a brand new CiteProc anyway
			// and so multiple resets here are not a problem because the raw cache would be empty anyway)
			CiteProc.ResetProcessorState();

			var result = new JSInlineCitation(CiteProc);

			if (idToUse != null)
			{
				result.CitationID = idToUse;
			}

			result.Properties.NoteIndex = 0;

			foreach(var itemSource in itemSources)
			{
				var inlineCitationItem = CiteProc.CreateJSInlineCitationItem(itemSource);

				result.CitationItems.Add(inlineCitationItem);
			}

			// We store this before Citeproc gets hold of it!
			result.FieldCodeJSON = CiteProc.ToJSON(result.JSObject, JSONWhitespace).Replace('\n', '\v') + FieldCodeSeparator;

			return result;
		}

		public bool ContainsBibliography
		{
			get
			{
				foreach(Field field in document.Fields)
				{
					var fieldText = field.Code.Text;
					if (fieldText.Contains(DocearMarker) && fieldText.Contains(CslBibliographyMarker))
					{
						return true;
					}
				}

				return false;
			}
		}

		IEnumerable<Field> EnumerateCSLFields()
		{
			foreach(Field field in document.Fields)
			{
				var fieldText = field.Code.Text;
				if (!fieldText.Contains(DocearMarker)) continue;

				yield return field;
			}
		}

		internal List<Field> GetCSLFields()
		{
			var result = new List<Field>();

			var fields = document.Fields;
			if (fields == null) return result; // Safety null check

			foreach(Field field in fields)
			{
				var code = field.Code;
				if (code == null) continue; // Safety null check

				var fieldText = code.Text;
				if (fieldText == null) continue; // Safety null check
				if (!fieldText.Contains(DocearMarker)) continue;
				if (!fieldText.Contains(CslCitationMarker) && !fieldText.Contains(CslBibliographyMarker)) continue;

				result.Add(field);
			}

			return result;
		}

		static bool IsDocearField(Field field)
		{
			return field.Code.Text.Contains(DocearMarker);
		}

		static bool IsCitationField(Field field)
		{
			var codeText = field.Code.Text;

			return codeText.Contains(DocearMarker) && codeText.Contains(CslCitationMarker);
		}

		public static bool IsBibliographyField(Field field)
		{
			var codeText = field.Code.Text;

			return codeText.Contains(DocearMarker) && codeText.Contains(CslBibliographyMarker);
		}

		public void MovePrevious()
		{
			var previousField = GetPreviousFieldBeforeRange(document.Application.Selection.Range);
			if (previousField == null) return;

			previousField.Select();
		}

		Field GetPreviousFieldBeforeRange(Range range)
		{
			var rangeStart = range.Start;

			for(var i = document.Fields.Count; i > 0; i--)
			{
				var field = document.Fields.Item(i);
				if (field.Type != WdFieldType.wdFieldAddin) continue;

				if (field.Result.Start < rangeStart && IsDocearField(field))
				{
					return field;
				}
			}

			return null;
		}

		Field GetNextFieldAfterRange(Range range)
		{
			var rangeEnd = range.End;

			foreach(Field field in document.Fields)
			{
				if (field.Type != WdFieldType.wdFieldAddin) continue;

				if (field.Result.End > rangeEnd && IsDocearField(field))
				{
					return field;
				}
			}

			return null;
		}

		public void MoveNext()
		{
			var nextField = GetNextFieldAfterRange(document.Application.Selection.Range);
			if (nextField == null) return;

			nextField.Select();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="fullRefresh">True to compare against database; False for style change</param>
		public int Refresh(bool fullRefresh)
		{
			try
			{
				isUpdating = true;

				ResetCiteProc();

				if (CiteProc == null)
				{
					MessageBox.Show("Docear4Word could not process the selected style.\r\nPlease check the style is valid or choose a different style.\r\n\r\nFeel free to contact us telling us about the problem (please include the style file when sending an email). http://www.docear.org/docear/contact/", "Docear4Word", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return -1;
				}

				return new CitationInserter(this).Refresh(fullRefresh);
			}
			finally
			{
				isUpdating = false;
			}
		}
		
		void ResetCiteProc()
		{
			if (citeProc == null) return;

			citeProc.Dispose();

			citeProc = null;
		}
	}

	public enum FieldMatchType
	{
		None,
		Inside,
		Partial,
		Wrap,
		Prior
	}

	public class FieldMatch
	{
		readonly Field field;
		readonly FieldMatchType matchType;
		readonly bool isBibliography;

		public FieldMatch(Field field, FieldMatchType matchType, bool isBibliography)
		{
			this.field = field;
			this.matchType = matchType;
			this.isBibliography = isBibliography;
		}

		public Field Field
		{
			get { return field; }
		}

		public FieldMatchType MatchType
		{
			get { return matchType; }
		}

		public bool IsBibliography
		{
			get { return isBibliography; }
		}
	}

	public class SelectionManager
	{
		readonly DocumentController documentController;
		readonly Range range;

		List<FieldMatch> fieldMatches;

		public SelectionManager(DocumentController documentController)
		{
			this.documentController = documentController;

			range = documentController.Document.Application.Selection.Range.Duplicate;
		}

		public List<FieldMatch> FieldMatches
		{
			get
			{
				if (fieldMatches == null)
				{
					fieldMatches = new List<FieldMatch>();

					var cslFields = documentController.GetCSLFields();
					if (cslFields.Count != 0)
					{

						var rangeStart = range.Start;
						var rangeEnd = range.End;

						foreach (var field in cslFields)
						{
							var fieldEnd = field.Result.End;
							if (fieldEnd < rangeStart) continue; // Field completely before selection

							Debug.Assert(field.Code.Start < field.Result.Start);
							var fieldStart = field.Code.Start;

							var isExactlyBeforeField = rangeEnd + 1 == fieldStart && rangeEnd == rangeStart;
							if (fieldStart > rangeEnd && // Field completely after selection
							    // But not immediately following a single point
							    !isExactlyBeforeField
								) break;

							FieldMatchType fieldMatchType;

							if (rangeStart < fieldStart)
							{
								fieldMatchType = isExactlyBeforeField
								                 	? FieldMatchType.Prior
								                 	: rangeEnd <= fieldEnd
								                 	  	? FieldMatchType.Partial
								                 	  	: FieldMatchType.Wrap;
							}
							else
							{
								fieldMatchType = rangeEnd <= fieldEnd
								                 	? FieldMatchType.Inside
								                 	: FieldMatchType.Partial;
							}

							if (fieldMatchType == FieldMatchType.None) continue;

							Debug.WriteLine(fieldMatchType);
							fieldMatches.Add(new FieldMatch(field, fieldMatchType, DocumentController.IsBibliographyField(field)));
						}
					}
				}

				return fieldMatches;
			}
		}

		public bool IsSingleMatchPrior
		{
			get { return FieldMatches.Count == 1 && FieldMatches[0].MatchType == FieldMatchType.Partial; }
		}

		public bool IsSingleCitation
		{
			get { return FieldMatches.Count == 1 && !FieldMatches[0].IsBibliography; }
		}

		public bool IsOneOrMoreCitations
		{
			get
			{
				foreach(var fieldMatch in FieldMatches)
				{
					if (fieldMatch.IsBibliography) return false;					
				}

				return fieldMatches.Count > 0;
			}
		}

		public bool IsRange
		{
			get { return range.Start != range.End; }
		}

		public Range Range
		{
			get { return range; }
		}
	}
}