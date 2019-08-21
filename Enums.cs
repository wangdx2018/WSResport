using System ;

namespace AFC.WorkStation.ExcelReport
{
	public enum GroupAlign
	{
		none = 0,
		vGroup = 1,
		hGroup = 2,
		always = 3
	}


	public enum MergeOption
	{
		never = 0,
		Up = 1,
		Left = 2,
	}

	[Flags]
	public enum InsertOption
	{
		never = 0,
		onfirst = 1,
		onLast = 2,
		BeforeChange = 4,
		afterChange = 8,
		always = 16
	}
	
	public enum  XlInsertShiftDirection
	{
		xlShiftToRight ,
		xlShiftDown
	}
	
	public enum  XlDeleteShiftDirection
	{
		xlShiftToLeft,
		xlShiftUp
	}
}