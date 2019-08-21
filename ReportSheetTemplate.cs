using System ;
using System.Collections ;
using System.Collections.Generic ;
using System.Data ;
using System.Data.SqlClient ;
using System.Drawing ;
using System.Text ;
using AFC.WorkStation.DB ;
using Spire.Xls ;

using Range=Spire.Xls.CellRange ;
namespace AFC.WorkStation.ExcelReport
{
	public class ReportSheetTemplate
	{
		
		public int startRowIndex ;
		
		public int rowCount = 0 ;
		
		public List<TplBlock> blockList = new List<TplBlock> ();
		
		public Dictionary<string, object> paramMap = new Dictionary<string, object> ();
		
		public DataSet dset = null ;
		public Dictionary<string, string> sqlList = new Dictionary<string, string>();
		public bool autoFit = false ;
		// public object [][] table = null ;

		public Worksheet sheet = null;
		public List<Rectangle> pics ;

		public void LoadDataSource (DBO db)
		{
			if (dset == null)
				dset = new DataSet ();
			
			/* Prepare & Execute SQL */
			foreach (KeyValuePair<string, string> pair in sqlList)
			{
				string tableName = pair.Key ;
				string sql = pair.Value ;
				
				sql = PrepareSql (sql, paramMap) ;
				
				// Execute SQL.
				int retcode ;
				DataSet sqlSet = db.SqlQuery (out retcode, sql) ;
				
				if (retcode != 0)
					throw new Exception ("Execute SQL Error: [" + sql + "]") ;

				DataTable table = sqlSet.Tables [0] ;
				table.TableName = tableName ;
				sqlSet.Tables.Remove (table);
				dset.Tables.Add(table);
				// dset.Merge(table);
			}
		}

		private string PrepareSql (string sql, Dictionary<string, object> map)
		{
			StringBuilder buf = new StringBuilder ();
			
			int index = -1 ;
			int startIndex = 0 ;
			while ((index = sql.IndexOf ('{', startIndex)) >= 0)
			{
				buf.Append (sql.Substring (startIndex, index - startIndex)) ;
				// find next matched '}'
				int nextIndex = sql.IndexOf ('}', index) ;
				
				if (nextIndex < 0)
				{
					// no match ?
					Console.WriteLine ("Warnning: param {} not matched.");
					startIndex = index + 1 ;
					continue ;
				}
				
				if (index < nextIndex - 1)
				{
					string pName = sql.Substring (index + 1, nextIndex - 1 - index) ;

					object pValue ;
					
					if (map.TryGetValue (pName, out pValue))
					{
						buf.Append (pValue) ;
					}
				}
				startIndex = nextIndex + 1;
			}
			
			if (startIndex < sql.Length - 1)
			{
				buf.Append (sql.Substring (startIndex)) ;
			}
			
			return buf.ToString () ;
		}

		public void FillTemplate ()
		{
			for (int i = 0; i < blockList.Count; i++)
			{
				TplBlock block = blockList [i] ;
				
				if (block.copyOnly)
				{
					Range range = RangeHelper.InsertCopyRange (sheet, block.tplRange,
					                                           block.tplColumCount, block.tplRowCount,
					                                           block.startColIndex, startRowIndex + rowCount,
					                                           XlInsertShiftDirection.xlShiftDown) ;

					IEnumerator e = RangeHelper.GetRangeCells(range);
					
					while (e.MoveNext ())
					{
						Range cell = (Range) e.Current ;
						
						if (cell.HasMerged)
							continue ;
						
						string s = cell.Value2 as string ;
						
						if (s != null && s.StartsWith ("#") && s.Length > 1 
						    && paramMap != null)
						{
							s = s.Substring (1) ;
							string[] s2 = s.Split (new char[] {':'}, StringSplitOptions.RemoveEmptyEntries) ;


							object pValue = null ;
							paramMap.TryGetValue (s2 [0], out pValue) ;
							string format = "" ;
							if (s2.Length > 1)
							{
								format = s2 [1].ToLower () ;
							}
							RangeHelper.UpdateCellValue (this, cell, pValue, format);
							
						}

						CellRange origin = RangeHelper.GetRange(sheet, cell.Column, cell.Row - startRowIndex - rowCount + block.startRowIndex, 1, 1);
						if (origin.HasMerged)
						{
							// doMerge
							int col = origin.MergeArea.Column ;
							int mWidth = origin.MergeArea.ColumnCount ;
							int row = origin.MergeArea.Row + startRowIndex + rowCount - block.startRowIndex ;
							int mHeight = origin.MergeArea.RowCount ;

							CellRange mRange = RangeHelper.GetRange (sheet, col, row, mWidth, mHeight) ;
							if (! mRange.HasMerged)
								mRange.Merge () ;	
						}
					} 
					
					rowCount += block.tplRowCount ;
					
				}
				else
				{
					
					
					block.startRowIndex = startRowIndex + rowCount ;

					if (block.isChart)
						// chart block should be filled at last.
						return;
					
					// check cloumn table first 
					if (! string.IsNullOrEmpty (block.tplColTableName) &&
						dset.Tables.Contains(block.tplColTableName))
					{
						block.CreateDColumns (dset.Tables [block.tplColTableName]) ;
					}
					
					if (! dset.Tables.Contains (block.tableName))
					{
						throw new ArgumentException ("DataTable [" + block.tableName + "] of Block [" + block.name + "] not found in DataSet!") ;
					}

					
                    if (dset.Tables[block.tableName].Rows.Count <= 0 && block.emptyTableName !=null)
				    {
                        rowCount += block.FillBlock(dset.Tables[block.emptyTableName]);
				    }
				    else
					    rowCount += block.FillBlock (dset.Tables [block.tableName]) ;
				}
				
				
			}
		}
		
		public void Clear ()
		{
			sheet.RemoveRange(RangeHelper.GetEntireCol (sheet, 1)) ;
			sheet.RemoveRange(RangeHelper.GetEntireCol(sheet, 1)) ; 
			//.Delete(XlDeleteShiftDirection.xlShiftToLeft);
			
			for (int i = 1 ; i < startRowIndex; i ++)
			{
				sheet.RemoveRange (RangeHelper.GetEntireRow(sheet, 1)) ; 
				// .Delete(XlDeleteShiftDirection.xlShiftUp);
			}
		}
	}
}