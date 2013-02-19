using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

using Docear4Word.BibTex;

using Word;

namespace Docear4Word
{
	public partial class DocumentController
	{
		class CitationInserter
		{
			readonly DocumentController documentController;
			readonly CiteProcRunner citeProc;
			readonly Document document;
			readonly List<JSCitationIDAndIndexPair> citationList;
			readonly List<Field> bibliographyWordFields = new List<Field>();
			readonly Dictionary<string, JSInlineCitation> inlineCitationCache;

			int suspendRedrawCount;

			public CitationInserter(DocumentController documentController)
			{
				this.documentController = documentController;
				citeProc = documentController.CiteProc;
				document = documentController.document;
				inlineCitationCache = documentController.inlineCitationCache;

				citationList = new List<JSCitationIDAndIndexPair>(documentController.document.Fields.Count);
			}

			Field InsertCitationField(Range range, JSInlineCitation citation)
			{
				var field = range.Fields.Add(range, WdFieldType.wdFieldQuote, InsertingCitationMessage, false);

				var fieldCodeText = CreateFieldCodeText(citation);

				field.Code.Text = fieldCodeText;

				SetCitation(fieldCodeText, citation);
	
				FormatCitationField(field);

				return field;
			}

			static void FormatCitationField(Field field)
			{
				var warningRange = field.Code.Duplicate;
				warningRange.Start = warningRange.Start + warningRange.Text.IndexOf(DoNotModifyWarningMessage);
				warningRange.End = warningRange.Start + DoNotModifyWarningMessage.Length;
				warningRange.Font.Bold = 1;

				var jsonRange = field.Code.Duplicate;
				jsonRange.Start = jsonRange.Start + Helper.GetCitationJSONStart(jsonRange.Text);
				jsonRange.Font = FixedWidthCodeFont;
				
				field.Code.NoProofing = 1;
				field.Locked = true;
			}

			public void EditCitation(Field field, JSInlineCitation citation)
			{

				try
				{
					SuspendRedraw();

					var fieldCodeText = CreateFieldCodeText(citation);

					field.Code.Text = fieldCodeText;

					// Update the cache
					RemoveCitation(fieldCodeText);
					SetCitation(fieldCodeText, citation);
	
					FormatCitationField(field);

					var jsCitations = Reset();

					var jsResult = citeProc.RestoreProcessorState(jsCitations);

					ApplyResult(jsResult);

					UpdateBibliographyFields();
				}
				finally
				{
					ResumeRedraw();
				}
			}

			public void InsertCitation(Range range, JSInlineCitation newCitation)
			{
				try
				{
					SuspendRedraw();

					// Insert a field into the document
					InsertCitationField(range, newCitation);

					var jsCitations = Reset();

					var jsResult = citeProc.RestoreProcessorState(jsCitations);

					ApplyResult(jsResult);

					UpdateBibliographyFields();
				}
				finally
				{
					ResumeRedraw();
				}
			}

			void SuspendRedraw()
			{
				suspendRedrawCount++;
				document.Application.ScreenUpdating = false;
			}

			void ResumeRedraw()
			{
				suspendRedrawCount--;

				if (suspendRedrawCount == 0)
				{
					document.Application.ScreenUpdating = true;
				}
			}

			public void InsertBibliography(Range range)
			{
				Reset();

				var field = range.Fields.Add(range, WdFieldType.wdFieldQuote, InsertingBibliographyMessage, false);
				bibliographyWordFields.Add(field);

				UpdateBibliographyFields();
			}

			void UpdateBibliographyFields()
			{
				if (bibliographyWordFields.Count == 0) return;

				var bibliographyResult = citeProc.MakeBibliography();

				try
				{
					SuspendRedraw();

					var formatter = bibliographyResult != null
					                	? new BibliographyRangeFormatter(bibliographyResult)
					                	: null;

					foreach (var bibliographyField in bibliographyWordFields)
					{
						bibliographyField.Code.Text = " " + AddInMarker + " " + DocearMarker + " " + CslBibliographyMarker;

						if (formatter != null)
						{
							formatter.CreateBibliography(bibliographyField.Result);
						}
						else
						{
							bibliographyField.Result.Text = String.Empty;
						}
					}
				}
				finally
				{
					ResumeRedraw();
				}
			}

			static string CreateFieldCodeText(JSInlineCitation citation)
			{
				// No point in putting in \n (which is what we had originally
				// since Word replaces them with \r and that buggers up
				// the cache lookups!
				var sb = new StringBuilder();
				sb.Append(AddInMarker);
				sb.Append(" ");
				sb.Append(DocearMarker);
				sb.Append(" ");
				sb.Append(CslCitationMarker);
				sb.Append(FieldCodeSeparator);
				sb.Append(DoNotModifyWarningMessage);
				sb.Append(FieldCodeSeparator);

				// This must be done BEFORE the citation is sent to citeproc
				// otherwise additional fields are added!!
				sb.Append(citation.FieldCodeJSON);

				return sb.ToString();
			}

			/// <summary>
			/// 
			/// </summary>
			/// <returns>A list of the JS citation objects.</returns>
			object[] Reset()
			{
				citationList.Clear();
				bibliographyWordFields.Clear();

				var result = new List<object>();

				foreach (var cslField in documentController.EnumerateCSLFields())
				{
					if (IsBibliographyField(cslField))
					{
						bibliographyWordFields.Add(cslField);
						continue;
					}

					if (!IsCitationField(cslField)) continue;

					var existingCitation = GetCitation(cslField.Code.Text);

					result.Add(existingCitation.JSObject);

					var citationIDAndIndexPair = new JSCitationIDAndIndexPair
					                            	{
					                            		ID = (string) existingCitation.CitationID,
					                            		FieldSource = cslField
					                            	};

					citationList.Add(citationIDAndIndexPair);

					for (var i = 0; i < existingCitation.CitationItems.Length; i++)
					{
						citeProc.CacheRawCitationItem(existingCitation.CitationItems[i].ItemData);
					}
				}

				return result.ToArray();
			}

			static bool ContainsUnknownItems(IEnumerable<EntryAndPagePair> sources)
			{
				foreach(var source in sources)
				{
					if (source.ID.StartsWith("_")) return true;
				}

				return false;
			}

			int UpdateCitationsFromDatabase(BibTexDatabase database)
			{
				if (database == null || !Settings.Instance.RefreshUpdatesCitationsFromDatabase) return -1;

				var changesMade = 0;

				foreach (var cslField in documentController.EnumerateCSLFields())
				{
					if (!IsCitationField(cslField)) continue;

					var existingCitation = GetCitation(cslField.Code.Text);

					var sources = Helper.ExtractSources(existingCitation, database);

					// We can't risk losing a whole citation cluster because one item is not in the database
					// so we skip the compare and keep the inline version
					if (sources.Count != existingCitation.CitationItems.Length) continue;

					// We also won't update a cluster containing unknown items
					// since this behaviour is undefined
					if (ContainsUnknownItems(sources)) continue;

					var databaseCitation = documentController.CreateInlineCitation(sources, existingCitation.CitationID);
					Debug.Assert(existingCitation.JSObject != databaseCitation.JSObject);

					if (DatabaseVersionIsDifferent(existingCitation, databaseCitation))
					{
						// Change the Field Code Text
						cslField.Code.Text = CreateFieldCodeText(databaseCitation);
						FormatCitationField(cslField);

						// Increment the changed counter
						changesMade++;
					}
				}

				return changesMade;
			}

			void ApplyResult(JSProcessCitationResult jsResult)
			{
				foreach (var update in jsResult.Items)
				{
					var entryToUpdate = citationList[update.Index];

					new RangeFormatter().AssignHtml(entryToUpdate.FieldSource.Result, update.String);
				}
			}

			public int Refresh(bool fullRefresh)
			{
				// Is this the best place to do this??
				ClearCitationsCache();

				var changesMade = fullRefresh
				              	? UpdateCitationsFromDatabase(documentController.GetDatabase())
				              	: -1;
				
				var jsCitations = Reset();

				var jsResult = citeProc.RestoreProcessorState(jsCitations);

				try
				{
					SuspendRedraw();

					ApplyResult(jsResult);

					UpdateBibliographyFields();
				}
				finally
				{
					ResumeRedraw();
				}

				return changesMade;
			}

			static bool DatabaseVersionIsDifferent(JSInlineCitation inlineCitation, JSInlineCitation databaseCitation)
			{
				return databaseCitation.FieldCodeJSON != inlineCitation.FieldCodeJSON;
			}

			void RemoveCitation(string fieldCodeText)
			{
				inlineCitationCache.Remove(fieldCodeText);
			}

			void SetCitation(string fieldCodeText, JSInlineCitation citation)
			{
				inlineCitationCache[fieldCodeText] = citation;
			}

			void ClearCitationsCache()
			{
				inlineCitationCache.Clear();
			}

			JSInlineCitation GetCitation(string fieldCodeText)
			{
				JSInlineCitation result;

				if (!inlineCitationCache.TryGetValue(fieldCodeText, out result))
				{
					var citationJSON = Helper.ExtractCitationJSON(fieldCodeText);

					result = JSInlineCitation.FromJSON(citeProc, citationJSON);

					SetCitation(fieldCodeText, result);
				}

				return result;
			}
		}
	}

	public class JSProcessCitationResult
	{
		readonly JSProcessCitationDataResult data;
		readonly JSProcessCitationIndexStringPair[] items;

		public JSProcessCitationResult(): this(null, null)
		{}

		public JSProcessCitationResult(JSProcessCitationDataResult data, JSProcessCitationIndexStringPair[] items)
		{
			this.data = data;
			this.items = items ?? new JSProcessCitationIndexStringPair[0];
		}

		public bool BibChange
		{
			get { return data.BibChange; }
		}

		public JSProcessCitationDataResult Data
		{
			get { return data; }
		}

		public JSProcessCitationIndexStringPair[] Items
		{
			get { return items; }
		}

		public class JSProcessCitationDataResult: JSObjectWrapper
		{
			const string BibChangeName = "bibchange";

			public JSProcessCitationDataResult(IJSContext context, object jsObject): base(context, jsObject)
			{
			}

			public bool BibChange
			{
				get { return (bool) GetProperty(BibChangeName); }
			}
		}
	}

	[ComVisible(false)]
	public class JSCitationIDAndIndexPair
	{
		public string ID { get; set; }
		public int Index { get; set; }
		public Field FieldSource { get; set; }
		public bool IsPre { get; set; }
		public bool IsPost { get; set; }
	}

	[ComVisible(false)]
	public class JSProcessCitationIndexStringPair
	{
		public int Index { get; set; }
		public string String { get; set; }
	}

}
