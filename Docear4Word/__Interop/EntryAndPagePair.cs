using System;

using Docear4Word.BibTex;

namespace Docear4Word
{
	public class EntryAndPagePair
	{
		readonly Entry entry;
		readonly string pageNumberOverride;
		readonly string id;

		public EntryAndPagePair(Entry entry, string pageNumberOverride = null)
		{
			this.entry = entry;

			// Ensure page number override is trimmed and, if empty, reset to null
			if (pageNumberOverride != null)
			{
				pageNumberOverride = pageNumberOverride.Trim();

				if (pageNumberOverride.Length == 0) pageNumberOverride = null;
			}

			this.pageNumberOverride = pageNumberOverride;

			id = this.pageNumberOverride == null
			     	? entry.Name
			     	: entry.Name + "#" + pageNumberOverride;
		}

		public Entry Entry
		{
			get { return entry; }
		}

		public string EntryName
		{
			get { return entry.Name; }
		}

		public string PageNumberOverride
		{
			get { return pageNumberOverride; }
		}

		// This is the one with a '#' if there is a PageNumberOverride
		public string ID
		{
			get { return id; }
		}
	}


}