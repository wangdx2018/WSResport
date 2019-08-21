using System ;
using System.Collections ;
using System.Collections.Generic ;
using System.Globalization ;
using System.IO ;
using System.Reflection ;
using System.Text ;

// using Microsoft.Office.Interop.Excel ;
using Spire.Xls ;
using Range=Spire.Xls.CellRange ;
using System.Data ;

namespace AFC.WorkStation.ExcelReport
{
	public class RangeHelper
	{
		public static int getColValueCount = 0 ;
		
	
		public static object GetColValue(DataTable table, int row, string colName)
		{
			getColValueCount ++ ;
			// return null ;
			
			// int colIndex ;
			int colIndex = table.Columns.IndexOf (colName) ;
			if (string.IsNullOrEmpty(colName) || colIndex < 0)
				return null;

			return table.Rows[row][colIndex];
		}
		
		public static object GetColValue(DataTable table, int row, int colIndex)
		{
			getColValueByIndexCount++;
			// return null ;

			// int colIndex ;
			// int colIndex = table.Columns.IndexOf(colName);
			if (/*string.IsNullOrEmpty(colName) || */colIndex < 0)
				return null;

			return table.Rows[row][colIndex];
		}
		public static object MakeCellName (int column, int row)
		{
			column -= 1 ;

			string value = GetExcelColIndex (column) ;

			return value + row ;
		}

		private static string GetExcelColIndex (int column)
		{
			string value = "" ;
			int deg = 0 ;
			do
			{
				value = "" + (char)((column % 26 - (deg > 0 ? 1 : 0)) + 'A') + value ;
				deg ++ ;
				column = column/26 ;
			} while (column > 0) ;
			
			return value ;
		}

		public static int getRangeCallTimes = 0 ;
		public static Range GetRange (Worksheet sheet, int startColumn, int startRow, int columns, int rows)
		{
			getRangeCallTimes ++ ;
			return sheet [startRow, startColumn, startRow + rows - 1, startColumn + columns - 1] ;
			/*return sheet.get_Range (MakeCellName (startColumn, startRow),
			                        MakeCellName (startColumn + columns - 1, startRow + rows - 1)) ;*/
		}

		public static Range GetLine (Worksheet sheet, int startColumn, int startRow)
		{
			return sheet [startRow, startColumn].EntireRow ;
			// return sheet.get_Range (MakeCellName (startColumn, startRow), Missing.Value).EntireRow ;
		}

		public static Range GetCell (Worksheet sheet, int startColumn, int startRow)
		{
			return GetRange (sheet, startColumn, startRow, 1, 1) ;
			// return sheet.get_Range (MakeCellName (startColumn, startRow), Missing.Value) ;
		}

		public static Range InsertCopyRange (Worksheet sheet, Range orign, int columns, int rows, int targetColumn,
		                                     int targetRow, XlInsertShiftDirection direction)
		{
			return InsertCopyRange (sheet, orign, columns, rows, targetColumn,
			                 targetRow, direction, 1) ;
		}
		
		public static int insertCopyRangeCallTimes = 0 ;
		public static Range InsertCopyRange (Worksheet sheet, Range orign, int columns, int rows, int targetColumn,
		                                     int targetRow, XlInsertShiftDirection direction, 
		                                     int lastColCount)
		{
			insertCopyRangeCallTimes++;
			// return GetRange(sheet, targetColumn, targetRow, columns, rows);
			int orgColumn = orign.Column ;
			int orgRow = orign.Row ;
			
			Range target = GetRange (sheet, targetColumn, targetRow, columns, rows) ;
			if (direction == XlInsertShiftDirection.xlShiftToRight)
			{
				// insert blank 
				// target.Insert (direction, Missing.Value) ;
				// Move target from origin to Right first 

				Range movedRange = GetRange(sheet, targetColumn, targetRow, lastColCount, rows);
				Range movedNextRange = GetRange(sheet, targetColumn + columns, targetRow, lastColCount, rows);
				movedRange.Move(movedNextRange, true, false);
				
				target = GetRange (sheet, targetColumn, targetRow, columns, rows) ;
				target.UnMerge ();
			}
			orign.Copy (target, false, true) ;

			switch(direction)
			{
			case XlInsertShiftDirection.xlShiftDown:
				// copy row height ;
				for (int i = 0 ; i < rows ; i ++)
				{
					Range sRow = GetEntireRow(sheet, i + orgRow);
					Range tRow = GetEntireRow(sheet, i + targetRow);
					try
					{
						tRow.RowHeight = sRow.RowHeight ;
					}
					catch (Exception e)
					{
						Console.WriteLine ("Set RowHeight Error: " + e);
					}
				}
					break ;
				case XlInsertShiftDirection.xlShiftToRight:
					// copy col width ;
					for (int i = 0; i < columns; i++)
					{
						Range sCol = GetEntireCol(sheet, i + orgColumn);
						Range tCol = GetEntireCol(sheet, i + targetColumn + 1);

						try
						{
							tCol.ColumnWidth = sCol.ColumnWidth;
						}catch (Exception e)
						{
							Console.WriteLine("Set ColumnWidth Error: " + e);
						}
					}
					break ;
			}
			
			
			return target ;
		}


		public static Range GetEntireRow (Worksheet sheet, int rowIndex)
		{
			return sheet [rowIndex, 1].EntireRow ;
			// return sheet.get_Range ("A"+rowIndex.ToString (), Missing.Value).EntireRow ;
		}

		public static Range GetEntireCol (Worksheet sheet, int colIndex)
		{
			return sheet[1, colIndex].EntireColumn;
			// return sheet.get_Range(GetExcelColIndex (colIndex - 1)+"1", Missing.Value).EntireColumn;
		}

		public static IEnumerator GetRangeCells(Range range)
		{
			Type t = range.GetType();

			PropertyInfo pCells = t.BaseType.GetProperty("Cells");

			Array cells = (Array)pCells.GetValue(range, new object[0]);
			return cells.GetEnumerator();

		}
		
		public static int MergeRanges (Range currentCellRange, MergeOption mOption)
		{
			Worksheet sheet = currentCellRange.Worksheet;
			Range range = null;
			// clear value first.
			currentCellRange.Value2 = null;
			currentCellRange.Text = "";
			
			switch (mOption)
			{
				case MergeOption.Up:
					
					// Select previous range.
					range = GetRange(sheet, currentCellRange.Column, currentCellRange.Row - 1, 1, 1) ;

					if (range.HasMerged)
					{
						Range marea = range.MergeArea ;
						int mWidth = currentCellRange.Column - marea.Column + 1 ;
						int mHeight = currentCellRange.Row - marea.Row + 1 ;

						/*if (mWidth <= marea.ColumnCount && mHeight <= marea.RowCount)
						{
							Console.WriteLine ("Has be merged. ignore");
							return 1 ;
						}*/
						
						if (mWidth < marea.ColumnCount)
							mWidth = marea.ColumnCount ;
						
						if (mHeight < marea.RowCount)
							mHeight = marea.RowCount ;
						
						range = GetRange(sheet,
		                              marea.Column, marea.Row,
		                              mWidth ,
		                              mHeight) ;
						/*
						Console.WriteLine("Up cell has been merged: [" +
										   marea.Column + ", " + marea.Row + "], " +
										   (currentCellRange.Column - marea.Column + 1) + ", " +
										   (currentCellRange.Row - marea.Row + 1) + ", [" +
										  marea.ColumnCount + "," + marea.RowCount + "]");
						*/
						range.Merge () ;
					}
					else
					{
						range = GetRange(sheet,
		                             currentCellRange.Column, 
		                             currentCellRange.Row - 1, 
		                             1, 2);
						range.Merge();
					}
					
					return 1;
				case MergeOption.Left:
					// merge with Left cell.
					
					/*
					Console.WriteLine("Merge Left: get left cell first: [" + 
					                  (currentCellRange.Column ) +
					                  ", " + currentCellRange.Row + "].") ;
					*/
					range = GetRange (sheet, currentCellRange.Column - 1, currentCellRange.Row, 1, 1) ;

					if (range.HasMerged)
					{
						Range marea = range.MergeArea;
						int mWidth = currentCellRange.Column - marea.Column + 1;
						int mHeight = currentCellRange.Row - marea.Row + 1;
						if (mWidth < marea.ColumnCount)
							mWidth = marea.ColumnCount;

						if (mHeight < marea.RowCount)
							mHeight = marea.RowCount;

						range = GetRange(sheet,
									  marea.Column, marea.Row,
									  mWidth,
									  mHeight);
						/*
						Console.WriteLine ("Left cell has been merged: [" +
						                   marea.Column + ", " + marea.Row + "], " +
						                   (currentCellRange.Column - marea.Column + 1) + ", " +
						                   (currentCellRange.Row - marea.Row + 1) + ", [" +
						                  marea.ColumnCount + "," + marea.RowCount + "]");
						*/
						range.Merge () ;
					}
					else
					{
						range = GetRange(sheet,
		                             currentCellRange.Column - 1,
		                             currentCellRange.Row,
		                             2, 1);
						
						range.Merge () ;
					}
					
					// range.Merge ();
					return 1;
				case MergeOption.never:
				default:
					return 0;
			}
		}

		public const string FORMAT_DATE = "date";
		public const string FORMAT_MONTH = "month";
		public const string FORMAT_DATE_TIME = "datetime";
		public const string FORMAT_TIME = "time";
		public const string FORMAT_FEN = "fen";
		public const string FORMAT_NUMBER = "num";
		public const string FORMAT_TIME_SPAN = "span";
	
		
		static int UpdateCellValueCallTimes = 0 ;
		public static int getColValueByIndexCount = 0 ;

		public static void UpdateCellValue (ReportSheetTemplate tpl, CellRange cell, object value, string format)
		{
			UpdateCellValueCallTimes ++ ;
			// return ;
			value = FormatValue (tpl, value, format) ;

			try
			{
				switch(format)
				{
					case FORMAT_DATE:
					case FORMAT_MONTH:
					case FORMAT_DATE_TIME:
					case FORMAT_TIME:
					case FORMAT_TIME_SPAN :
						// cell.DateTimeValue = (DateTime) value ;
						// return ;
						break ;
					case FORMAT_FEN:
					case FORMAT_NUMBER:
				        try
				        {
				            cell.NumberValue = Convert.ToDouble(value) ;
				        }
				        catch (Exception e)
				        {
                            // ignore exception.
                            cell.NumberValue = 0;
				        }
						return ;
					default :
						break ;
				}
			}
			catch (Exception e)
			{
				// ignore exception.
			}


			cell.Text = Convert.ToString(value);
		}

		public static object FormatValue(ReportSheetTemplate tpl, object value, string format)
		{
			if (string.IsNullOrEmpty(format) || value == null)
				return value;

			switch (format)
			{
				case FORMAT_DATE:
					return ValueToDate(value);
				case FORMAT_MONTH:
					return ValueToMonth(value);
				case FORMAT_TIME:
					return ValueToTime(value);
				case FORMAT_DATE_TIME:
					return ValueToDateTime(value);
				case FORMAT_TIME_SPAN :
					return ValueToTimeSpan (value, tpl) ;
				case FORMAT_FEN:
					return ValueToFen(value);
				default:
					return value;
			}
		}

		private static object ValueToTimeSpan (object value, ReportSheetTemplate tpl)
		{
			int spanNumber ;
			object span ;
			DateTime time ;
			if (value == null || 
			    ! tpl.paramMap.TryGetValue ("time_span", out span) ||
				! int.TryParse(span.ToString (), out spanNumber) || 
				! DateTime.TryParseExact (value.ToString (), "yyyyMMddHHmmss",
					null, DateTimeStyles.None, out time))
				return ValueToTime (value) ;
			
			DateTime time2 = time.Add (new TimeSpan (0, 0, spanNumber, 0)) ;
			
			return time.ToString ("HH:mm") + "-" + time2.ToString ("HH:mm") ;
		}

		private static object ValueToMonth (object value)
		{
			string str = value.ToString();

			if (str.Length < 6)
				return value;
			if (str.Length > 6)
				str = str.Substring(0, 6);
			try
			{
				return DateTime.ParseExact(str, "yyyyMM", null).ToString("yyyy年MM月"); // .ToString();
			}
			catch (Exception e)
			{
				Console.WriteLine("parse Month Error [" + str + "]: " + e);
				return value;
			}
		}

		private static object ValueToFen(object value)
		{
			try
			{
				double d = Convert.ToDouble(value);
				return d / 100;
			}
			catch (Exception e)
			{
				Console.WriteLine("parse double Error [" + value + "]: " + e);

				return value;
			}
		}

		private static object ValueToDateTime(object value)
		{
			string str = value.ToString();

			if (str.Length < 14)
				return value;
			if (str.Length > 14)
				str = str.Substring(0, 14);
			try
			{
				return DateTime.ParseExact(str, "yyyyMMddHHmmss", null).ToString("yyyy年MM月dd日 HH:mm:ss"); // .ToString();
			}
			catch (Exception e)
			{
				Console.WriteLine("parse Datetime Error [" + str + "]: " + e);
				return value;
			}
		}

		private static object ValueToTime(object value)
		{
			string str = value.ToString();

			if (str.Length < 6)
				return value;
			if (str.Length > 6)
				str = str.Substring(str.Length - 6);

			try
			{
				return DateTime.ParseExact(str, "HHmmss", null).ToString ("HH:mm:ss") ; // .ToShortTimeString();
			}
			catch (Exception e)
			{
				Console.WriteLine("parse Time Error [" + str + "]: " + e);
				return value;
			}
		}

		private static object ValueToDate(object value)
		{
			string str = value.ToString();

			if (str.Length < 8)
				return value;
			if (str.Length > 8)
				str = str.Substring(0, 8);
			try
			{
				return DateTime.ParseExact (str, "yyyyMMdd", null).ToString ("yyyy年MM月dd日") ;
			}
			catch (Exception e)
			{
				Console.WriteLine("parse Date Error [" + str + "]: " + e);
				return value;
			}

		}

		public static object GetColValueByIndex (DataTable table, int rowIndex, int colIndex)
		{
			getColValueByIndexCount++;
			// return null ;

			// int colIndex ;
			
			return table.Rows[rowIndex][colIndex];
		}
	}

	


	/* 
	 * Line template A,1 
	 * repeat from A, 10 
	 * cell (A,1), (B,1), (C,1), (D1)
	 * 
	 * repeat Cell : (A , 10 + repeatCount) 
	 * Merge count :  = 2 ;
	 * Merge Area  : (A, 10 + repeatCount 
	*/

	
	public class GroupDataHolder
	{
		
		public Dictionary<GroupValueSearchKey, object> valueList = new Dictionary<GroupValueSearchKey, object> (640*1024);
		
		public object GetValue (GroupValueSearchKey gkey, object defaultValue)
		{
			object value ;
			if (valueList.TryGetValue (gkey, out value))
				return value ;
			else 
				return defaultValue ;
		}
		
		public void AddValue (Dictionary<GroupValueSearchKey, bool> map, GroupValueSearchKey gkey, 
		                      System.Data.DataTable table, int valueIndex)
		{
			if (map.ContainsKey(gkey))
				// ignored.
				return ;

			// judge Data is Match 
			if (gkey.key != null)
			{
				SearchKey tmpKey = gkey.key ;
				while(tmpKey != null)
				{	
					if (tmpKey.rKey == null)
					{
						tmpKey.rKey = ReusedKey.FindReusedKey (tmpKey) ;
					}
					
					if (tmpKey.isFixedValue)
					{
						// Compare Fixed value Key only 
						object keyValue = null ;
					
						keyValue = tmpKey.rKey.GetReusedValue (table, valueIndex) ;
						// 	RangeHelper.GetColValue (table, valueIndex, tmpKey.colName) ;
						
						// Not match, return .
						if (! Equals (keyValue, tmpKey.keyValue))
							return ;
					}
					
					tmpKey = tmpKey.nextKey ;
					
				}
				/*SearchKey copyKey = gkey.key.Copy0 (false) ;
				copyKey.FillKey (table, valueIndex);
			
				// not match. 
				if (! copyKey.Equals (gkey.key))
					return ;*/
			}
			
			object lastValue ;
			if (!valueList.TryGetValue(gkey, out lastValue))
			{
				lastValue = 0 ;
				valueList.Add(gkey, lastValue);
			}

			if (gkey.rKey == null)
			{
				SearchKey key = new SearchKey ();
				key.colName = gkey.valueColName ;
				gkey.rKey = ReusedKey.FindReusedKey(key);
			}
			
			valueList [gkey] = CalculateForumla (gkey.formula, lastValue, 
			                                     gkey.rKey.GetReusedValue (table, valueIndex)) ;

			map[gkey/*.Copy()*/] = true;
		}

		public object CalculateForumla(string formulaName, object lastValue, object value)
		{
			formulaName = formulaName.ToLower();
			
			lastValue = CellForumla.Calculate(formulaName, lastValue, value);

			return lastValue;
		}
		
		private static SearchKey CreateTopKey (SearchKey key)
		{
			SearchKey topKey = key.Copy () ;
			topKey.keyValue = "" ;
			topKey.colName = "" ;
			topKey.nextKey = key ;
			topKey.isFixedValue = true ;
			topKey.keyValue = null ;
			return topKey ;
		}

		/*public void AddValue0(bool searchInChildren, SearchKey key, string valueColName, string formulaName, object value)
		{
			if (! searchInChildren)
			{
				GroupValue gValue ;
				if (! valueList.TryGetValue (valueColName, out gValue))
				{
					// not found, Add New Value.
					gValue = new GroupValue ();
					gValue.colName = valueColName ;
			
					valueList.Add (valueColName, gValue);
				}

				gValue.CalculateForumla(formulaName, value);
				
				key = key.nextKey ;
			}
			
			if (key != null)
			{
				// search next key
				GroupDataHolder nextDataHolder ;
				if (! children.TryGetValue (key, out nextDataHolder))
				{
					// create new holder 
					nextDataHolder = new GroupDataHolder ();
					nextDataHolder.colIndex = key.colIndex ; 
					nextDataHolder.colName = key.colName ; 
					nextDataHolder.keyValue = key.keyValue ;
					children.Add (key, nextDataHolder);
				}

				nextDataHolder.AddValue0(false, key, valueColName, formulaName, value);

			}
			
			
		}*/
	}
	
	public class SearchKey : KeyValuePair
	{
		public bool isFixedValue = false ;
		public SearchKey nextKey = null ;
		
		public ReusedKey rKey = null ;
		
		public void FillKey (System.Data.DataTable table, int rowIndex)
		{
			if (! isFixedValue)
			{
				// use context key value ;
				
				if (string.IsNullOrEmpty(colName))
					keyValue = null ;

				if (rKey == null)
					rKey = ReusedKey.FindReusedKey (this) ;
				
				keyValue = rKey.GetReusedValue (table, rowIndex) ;
					// RangeHelper.GetColValue (table, rowIndex, colName) ;
			}
			
			if (nextKey != null)
				nextKey.FillKey (table, rowIndex);
		}
		
		public SearchKey Copy ()
		{
			return Copy0 (true) ;
		}

		public SearchKey Copy0(bool copyValue)
		{
			SearchKey key = new SearchKey();
			key.colIndex = colIndex;
			key.colName = colName;
			key.rKey = rKey ;
			if (copyValue)
			{
				key.keyValue = keyValue;
				key.isFixedValue = isFixedValue;
			}

			if (nextKey != null)
				key.nextKey = nextKey.Copy0(copyValue);

			return key;
		}


		public override int GetHashCode ()
		{
			return base.GetHashCode () + 29*(nextKey != null ? nextKey.GetHashCode () : 0) ;
		}

		public override bool Equals (object obj)
		{
			if (this == obj) return true ;
			SearchKey searchKey = obj as SearchKey ;
			if (searchKey == null) return false ;
			if (!base.Equals (obj)) return false ;
			if (!Equals (nextKey, searchKey.nextKey)) return false ;
			return true ;
		}

		public override string ToString ()
		{
			return "SearchKey: [" + colName + ", " + keyValue + ", " + isFixedValue + "]" +
			       "next->" + nextKey ;
		}
	}
	
	/*public class GroupValue : KeyValuePair
	{
		private Dictionary<string, object> fValueList = new Dictionary<string, object>();
		public object GetFormulaValue (string formula)
		{
			object value ;
			if (fValueList.TryGetValue (formula, out value))
				return value ;
			
			throw new ArgumentOutOfRangeException ("formula", formula, "Formula Value not found in valueList.") ;
		}

		public object CalculateForumla (string formulaName, object value)
		{
			object lastValue ;
			formulaName = formulaName.ToLower () ;
			if (! fValueList.TryGetValue (formulaName, out lastValue))
			{
				lastValue = 0 ;
				fValueList.Add (formulaName, lastValue);
			}
			lastValue = CellForumla.Calculate(formulaName, lastValue, value);
			
			fValueList [formulaName] = lastValue ;
			
			return lastValue ;
		}

		
	}
	*/
	public class KeyValuePair
	{
		public const string COL_NAME_NONE = "";
		public const int COL_INDEX_NONE = -1;
		
		public int colIndex;
		public string colName;
		public object keyValue;
		
		public override int GetHashCode ()
		{
			return (colName != null ? colName.GetHashCode () : 0) + 29*(keyValue != null ? keyValue.GetHashCode () : 0) ;
		}

		public override bool Equals (object obj)
		{
			if (this == obj) return true ;
			KeyValuePair keyValuePair = obj as KeyValuePair ;
			if (keyValuePair == null) return false ;
			if (!Equals (colName, keyValuePair.colName)) return false ;
			if (!Equals (keyValue, keyValuePair.keyValue)) return false ;
			return true ;
		}
		
		public static bool CompareValue (object v1, object v2)
		{
			
			if (Equals (v1, v2))
				return true ;
			
			return string.Equals ("" + v1, "" + v2) ;
		}
	}
	
	public class ReusedKey : KeyValuePair
	{
		public DataTable lastUsedTable = null ;
		public int lastUsedColIndex = -1 ;
		public int lastUsedRowIndex = -1 ;
		public object lastValue ;
		
		public static int getReusedValueCallCount = 0 ;
		public static int columnNameFindCount = 0 ;
		public static int valueReusedCount = 0 ;
		public static List<ReusedKey> keyPool = new List<ReusedKey> ();
		
		public object GetReusedValue (DataTable table, int rowIndex)
		{
			getReusedValueCallCount ++ ;
			
			if (lastUsedTable != table)
			{
				lastUsedTable = table ;
				columnNameFindCount ++ ;
				lastUsedColIndex = table.Columns.IndexOf (colName) ;
				lastUsedRowIndex = -1 ;
				lastValue = null ;
			}
			else
			if (lastUsedRowIndex == rowIndex)
			{
				valueReusedCount ++ ;
				return lastValue ;
			}
			
			lastUsedRowIndex = rowIndex ;
			
			/*
			if (lastUsedColIndex < 0)
				lastValue = null ;
			else
				lastValue = table.Rows[rowIndex][lastUsedColIndex];
			*/
			lastValue = RangeHelper.GetColValue (table, rowIndex, lastUsedColIndex) ;

			return lastValue ;
			
		}
		
		public static ReusedKey FindReusedKey (SearchKey key)
		{
			for (int i = 0; i < keyPool.Count; i++)
			{
				ReusedKey reusedKey = keyPool [i] ;
				
				if (Equals (reusedKey.colName, key.colName))
					return reusedKey ;
			}

			ReusedKey rKey = new ReusedKey () ;
			rKey.colName = key.colName ;
			
			keyPool.Add (rKey);
			
			return rKey ;
		}
	}
	
	public class GroupValueSearchKey
	{
		public SearchKey key ;
		public string valueColName ;
		public string formula ;
		public ReusedKey rKey ;

		public GroupValueSearchKey Copy ()
		{
			GroupValueSearchKey gkey = new GroupValueSearchKey ();
			if (key != null)	
				gkey.key = key.Copy () ;
			
			gkey.formula = formula ;
			gkey.valueColName = valueColName;
			gkey.rKey = rKey ;
			return gkey;
		}

		public override int GetHashCode ()
		{
			int result = key != null ? key.GetHashCode () : 0 ;
			result = 29*result + (valueColName != null ? valueColName.GetHashCode () : 0) ;
			result = 29*result + (formula != null ? formula.GetHashCode () : 0) ;
			return result ;
		}

		public override bool Equals (object obj)
		{
			if (this == obj) return true ;
			GroupValueSearchKey groupValueSearchKey = obj as GroupValueSearchKey ;
			if (groupValueSearchKey == null) return false ;
			if (!Equals (key, groupValueSearchKey.key)) return false ;
			if (!Equals (valueColName, groupValueSearchKey.valueColName)) return false ;
			if (!Equals (formula, groupValueSearchKey.formula)) return false ;
			return true ;
		}

		public override string ToString ()
		{
			return "GroupValueSearchKey [" + formula + ", " + valueColName + "]" +
			       key ;
		}
	}
	
	public class CellForumla
	{
		public string formulaName = string.Empty;
		public List<GroupValueSearchKey> keyList = new List<GroupValueSearchKey> ();
			
		public object GetValue (GroupDataHolder holder, System.Data.DataTable table, int rowIndex)
		{
			object lastValue = 0 ;
			for (int i = 0; i < keyList.Count; i++)
			{
				GroupValueSearchKey gkey = keyList [i] ;
				if (gkey.key != null)
					gkey.key.FillKey (table, rowIndex);

				object value = holder.GetValue (gkey, null) ;
				/* 
				Console.WriteLine ("row ["+ rowIndex + "] Get value for gkey: " + gkey );
				Console.WriteLine ("row [" + rowIndex + "] = [" + value + "]") ;
				 */
				lastValue = Calculate (formulaName, lastValue, value) ;
			}
			
			return lastValue ;
		
		}

		public static object Calculate (string formulaName, object lastValue, object value)
		{
			switch (formulaName)
			{
				case "" :
					return value ;
				case "sum":
					lastValue = ToNumber(lastValue) + ToNumber(value);
					break;
				case "count":
					lastValue = ToNumber(lastValue) + 1;
					break;
				default:
					throw new ArgumentOutOfRangeException("formula", formulaName, "Formula [" + formulaName + "] is not supported.");
			}
			return lastValue ;
		}

		private static double ToNumber(object value)
		{
			if (value == null)
				return 0D;

			string svalue = value.ToString();
			double d = 0;
			double.TryParse(svalue, out d);
			return d;
		}

		public CellForumla Copy ()
		{
			CellForumla f = new CellForumla ();
			f.formulaName = formulaName ;
			for (int i = 0; i < keyList.Count; i++)
			{
				GroupValueSearchKey gkey = keyList [i] ;
				f.keyList.Add (gkey.Copy ());
			}
			
			return f ;
		}

		public override string ToString ()
		{
			StringBuilder buf = new StringBuilder ();
			buf.Append ("CellForumla [" + formulaName + "] keyCount = " + keyList.Count + "\n") ;
			for (int i = 0; i < keyList.Count; i++)
			{
				GroupValueSearchKey key = keyList [i] ;
				
				buf.Append ("\t") ;
				buf.Append (key.ToString ()) ;
				buf.Append ("\n") ;
			}
			
			
			return buf.ToString () ;
			
		}
		
		
	}
    
    //provide empty table
    public class EmptyTable
    {
        //provide a empty table with 1 empty row by default 
        public static DataTable  GetEmptyTalbe(DataTable sourceTable)
        {
           return GetEmptyTalbe(sourceTable, 1);
        }

        //provide a empty table with custum number of enmpty rows
        public static DataTable GetEmptyTalbe(DataTable sourceTable,int rowsCount)
        {
            if (rowsCount <= 0)
                return sourceTable;
            
            //DataRow row = sourceTable.NewRow();
            //row[0] = DBNull.Value;
            for (int i = 0; i < rowsCount;i++ )
            {
                DataRow row = sourceTable.NewRow();
                row[0] = DBNull.Value;
                sourceTable.Rows.Add(row); 
            }
            //DataTable table = new DataTable();
            //table.TableName = sourceTable.TableName;
            //table.Columns.Add(new DataColumn("emptyColum"));

            //table.Rows.Add(row);
            sourceTable.ExtendedProperties.Add("TableType", "CustumEmpty");
            return sourceTable;
        }
    }
	
	/*
	public class DataTree
	{
		public DataTree parent ;
		public List<DataTree> children ;
		public Dictionary<string ,object> valueList;
		public string keyName ;
		public object keyValue ;
		public bool isRoot ;
		
		public void AddGroupedKey (GroupValueSearchKey gKey)
		{
			SearchKey key = gKey.key ;
			
			while (key != null)
			{
				if (children == null)
					children = new List<DataTree> ();
				
				for (int i = 0; i < children.Count; i++)
				{
					DataTree child = children [i] ;
					if (child.keyName.Equals (key.colName) && child.keyValue)
						child.AddGroupedKey ();
				}
				
				key = key.nextKey ;
			}
		}
	}
	*/
	
}
