namespace Docear4Word
{
	public class JSIDandNotePair: JSObjectWrapper
	{
		public JSIDandNotePair(IJSContext context, string id, int index): base(context)
		{}

		JSTypedArray<JSNameVariable> Internal
		{
			get { return GetTypedArray<JSNameVariable>(CSLNames.Editor); }
		}

		 
	}
}