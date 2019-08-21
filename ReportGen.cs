using System;
using System.Collections ;
using System.Collections.Generic;
using System.Data ;
using System.Diagnostics ;
using System.Drawing ;
using System.IO ;
using System.Reflection ;
using AFC.WorkStation.DB ;
using AFC.WorkStation.ExcelReport ;
using Microsoft.Office.Interop.Excel ;
using Spire.Xls ;
using Rectangle=System.Drawing.Rectangle;
using Workbook=Spire.Xls.Workbook;
using Worksheet=Microsoft.Office.Interop.Excel.Worksheet;

namespace AFC.WorkStation.ExcelReport
{
	public class ReportGen
	{
    	public Application xlapp = null; 
		
		public DataSet set = new DataSet ();
		Dictionary<string,object> paramMap = new Dictionary<string, object> ();
		
		public DBO db ;
		
		/* 
		public void AddTable (DBO db, string sql, string tableName)
		{
			int retcode ;
			DataSet dset = db.SqlQuery (out retcode, sql) ;

			DataTable table = dset.Tables [0] ;
			if (! string.IsNullOrEmpty(tableName))
				table.TableName = tableName ;
			set.Merge(table);
			// set.Tables.Add () ;
		}
		*/
		
		public void AddParam (string name, object value)
		{
			paramMap [name] = value ;
		}
		
		public void LoadReport (string fileName)
		{
			LoadReport (fileName, false);
		}
		public void LoadReport (string fileName, bool autoFit)
		{
			Workbook book = new Workbook ();
			
		
			book.LoadFromFile (fileName);
			
			ReportBookTemplate tplbook = new ReportBookTemplate ();
			tplbook.LoadTemplate (db, book, paramMap);
			tplbook.FillTemplate ();
			
			List<ReportSheetTemplate> tplList = tplbook.sheetList ;
			// tpl.Clear();
		
			// Open with Excel

/*			xlapp.Visible = false;*/
			
			try
			{
				ClearReport (autoFit, book, tplList) ;
				// xlapp.Workbooks.Close ();
				
				// Copy Image 
				
			}
			finally 
			{
				book.Save();
				book.Dispose();
				
/*				xlapp.DisplayAlerts = true ;
			
				xlapp.Visible = true;
 */
				RemoveWarning(fileName);

				Process.Start (fileName) ;
			}


			// remove warnning sheet.
			
		}

		private void RemoveWarning (string fileName)
		{
			FileStream fp = null ;

			try
			{
				fp = new FileStream (fileName, FileMode.Open, FileAccess.ReadWrite) ;

				Biff8Helper.RemoveWarning (fp) ;
			}
			catch (Exception e)
			{
				Console.WriteLine("RemoveWarning Error: "  + e);
			}
			finally
			{
				if (fp != null)
					fp.Close ();
			}
			
		}

		private void ClearReport (bool autoFit, Workbook book, List<ReportSheetTemplate> tplList)
		{
			
			for (int i = 0; i < book.Worksheets.Count && 
			                i < tplList.Count; i++)
			{
				Spire.Xls.Worksheet worksheet = book.Worksheets [i] ;

				ReportSheetTemplate tpl = tplList [i] ;
				JoinTable(worksheet, tpl);
				// Clear Data 
				Clear(worksheet, tpl.startRowIndex);
					
				if (autoFit || tpl.autoFit)
				{
					CellRange range = RangeHelper.GetRange(worksheet, 1, 15, 50,100);
					range.AutoFitColumns ();

					
					/*for (int j = 1; j < 100; j++)
					{
						try
						{
							worksheet.AutoFitColumn (j);
						}
						catch (Exception e)
						{
							Console.Write (e) ;
						}
					}
					*/
				}
				
				// .GetType ().GetMethod ("AutoFit").Invoke (range, new object[0]) ;
				// copy image 
				/*for (int j = 0; tpl.pics != null && 
				                j < tpl.pics.Count && 
				                j < worksheet.Pictures.Count; j++)
				{
					Rectangle pic = tpl.pics [i] ;
					
					
					int tmp = worksheet.Pictures [i].TopRow ;
					tmp = worksheet.Pictures[i].TopRowOffset;
					tmp = worksheet.Pictures[i].LeftColumn;
					tmp = worksheet.Pictures[i].LeftColumnOffset;
					tmp = worksheet.Pictures [i].BottomRow ;
					tmp = worksheet.Pictures [i].BottomRowOffset ;
					tmp = worksheet.Pictures [i].RightColumn ;
					tmp = worksheet.Pictures [i].RightColumnOffset ;
					
					worksheet.Pictures [i].Left = pic.X ;
					worksheet.Pictures [i].Top = pic.Y ;
					worksheet.Pictures [i].Height = pic.Height ;
					worksheet.Pictures [i].Width = pic.Width ;
					
				}*/
				
			}
				
			// remove warnning sheet.
			
/*
			IEnumerator e = xlapp.ActiveWorkbook.Worksheets.GetEnumerator();
			while (e.MoveNext())
			{
				Worksheet sheet = (Worksheet)e.Current;

				if (sheet.Name.IndexOf("Warning") >= 0)
					sheet.Delete();
			}
			((_Worksheet)xlapp.ActiveWorkbook.Worksheets[1]).Activate ();
			((Worksheet)xlapp.ActiveWorkbook.Worksheets [1]).get_Range("A1", "A1").Activate();
			// only save activeWorkBook 
			xlapp.ActiveWorkbook.Save ();
			// only close activeWorkbook ;
			xlapp.ActiveWorkbook.Close (true, Missing.Value, Missing.Value);
			// xlapp.Workbooks.Close ();
 */
		}
		/* 
		private void ClearExcelReport(bool autoFit, string fileName, List<ReportSheetTemplate> tplList)
		{
			bool reusedFlag = false;
			if (xlapp.Workbooks.Count > 0)
			{
				reusedFlag = true;
			}
			xlapp.Workbooks.Open(fileName,
								  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
								  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
								  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
								  Missing.Value, Missing.Value);

			xlapp.DisplayAlerts = false;

			for (int i = 0; i < xlapp.ActiveWorkbook.Worksheets.Count &&
							i < tplList.Count; i++)
			{
				Worksheet worksheet = (Worksheet)xlapp.ActiveWorkbook.Worksheets[i + 1];

				ReportSheetTemplate tpl = tplList[i];
				JoinTable(worksheet, tpl);
				// Clear Data 
				Clear(worksheet, tpl.startRowIndex);

				if (autoFit || tpl.autoFit)
				{
					Range range = worksheet.get_Range("A1", "DZ1").EntireColumn;
					range.AutoFit();
				}
				// .GetType ().GetMethod ("AutoFit").Invoke (range, new object[0]) ;

			}

			// remove warnning sheet.

			IEnumerator e = xlapp.ActiveWorkbook.Worksheets.GetEnumerator();
			while (e.MoveNext())
			{
				Worksheet sheet = (Worksheet)e.Current;

				if (sheet.Name.IndexOf("Warning") >= 0)
					sheet.Delete();
			}
			((_Worksheet)xlapp.ActiveWorkbook.Worksheets[1]).Activate();
			((Worksheet)xlapp.ActiveWorkbook.Worksheets[1]).get_Range("A1", "A1").Activate();
			// only save activeWorkBook 
			xlapp.ActiveWorkbook.Save();
			// only close activeWorkbook ;
			xlapp.ActiveWorkbook.Close(true, Missing.Value, Missing.Value);
			// xlapp.Workbooks.Close ();
		}
		*/
		public static void JoinTable(Spire.Xls.Worksheet sheet, ReportSheetTemplate tpl)
		{
			if (tpl.blockList.Count < 2)
				return ;
			
			TplBlock firstBlock = tpl.blockList [1] ;
			int blockRow = firstBlock.startRowIndex ;
			int blocklastColum = firstBlock.startColIndex + firstBlock.colCount
			                     /* - 
			                     (firstBlock.dColumn == null ? 0 : firstBlock.dColumn.gCols)*/ ;
			
			int joinedRows = 0 ;
			
			for (int i = 2; i < tpl.blockList.Count; i++)
			{
				TplBlock block = tpl.blockList [i] ;
				
				if (block.joinat >= 0 && block.rowCount > 0)
				{
					// CopyRangeToFirstTable 
					CellRange range = RangeHelper.GetRange(sheet, block.startColIndex + block.joinat + 1, 
						block.startRowIndex - joinedRows,
						block.colCount,block.rowCount) ;
					range.Copy (
						RangeHelper.GetRange(sheet, blocklastColum + 1, firstBlock.startRowIndex, block.colCount, block.rowCount));
					
					
					/* range.EntireRow.Delete (Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp) ; */
					// delete rows.
					for (int k = 0 ; k < block.rowCount; k ++)
					{
						sheet.DeleteRow(block.startRowIndex - joinedRows);
					}
					
					joinedRows += block.rowCount ;
					if (block.dColumn != null && block.dColumn.startCellIndex == block.joinat)
					{
						// Merge Joined Table Columns.
						for (int j = 0; j < block.lineList.Count; j++)
						{
							TplLine line = block.lineList [j] ;
						
							if (!line.containsHGroup)
								continue ;

							Boolean hasMerged = false ;
							for (int k = 0; k < line.insertedRowList.Count; k++)
							{
								int rowIndex = line.insertedRowList [k] ;
								rowIndex = rowIndex - block.startRowIndex + blockRow ;

								CellRange leftRange = RangeHelper.GetRange (sheet, blocklastColum, rowIndex, 1, 1) ;
								if (leftRange.MergeArea != null)
									leftRange = RangeHelper.GetRange(sheet, leftRange.MergeArea.Column, leftRange.MergeArea.Row, 1, 1);
								CellRange rightRange = RangeHelper.GetRange(sheet, blocklastColum + 1, rowIndex, 1, 1);
								if (rightRange.MergeArea != null)
									rightRange = RangeHelper.GetRange(sheet, rightRange.MergeArea.Column, rightRange.MergeArea.Row, 1, 1);
							
								if (leftRange.Text.Equals (rightRange.Text))
								{
									
									// Merge 

									RangeHelper.GetRange(sheet, leftRange.Column, 
									              leftRange.Row,
									              rightRange.Column + rightRange.Columns.Length - leftRange.Column, 
								               
									              Math.Min(rightRange.Rows.Length, leftRange.Rows.Length) 
										).Merge () ;
									
									hasMerged = true ;
								}
							}
						
							if (! hasMerged)
								break ;
						} // end for 
					} // end if 
					blocklastColum += block.colCount - block.joinat ;
					
				}
			}
		}
		
		/* 
		public static void JoinTable(Worksheet sheet, ReportSheetTemplate tpl)
		{
			if (tpl.blockList.Count < 2)
				return ;
			
			TplBlock firstBlock = tpl.blockList [1] ;
			int blockRow = firstBlock.startRowIndex ;
			int blocklastColum = firstBlock.startColIndex + firstBlock.colCount
			                     //  - 
			                     // (firstBlock.dColumn == null ? 0 : firstBlock.dColumn.gCols)
									;
			
			int joinedRows = 0 ;
			
			for (int i = 2; i < tpl.blockList.Count; i++)
			{
				TplBlock block = tpl.blockList [i] ;
				
				if (block.joinat >= 0 && block.rowCount > 0)
				{
					// CopyRangeToFirstTable 
					Range range = GetExcelRange(sheet, block.startColIndex + block.joinat + 1, 
						block.startRowIndex - joinedRows,
						block.colCount,block.rowCount) ;
					range.Copy (
						GetExcelRange(sheet, blocklastColum + 1, firstBlock.startRowIndex, block.colCount, block.rowCount));
					
					
					range.EntireRow.Delete (Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp) ;
					joinedRows += block.rowCount ;
					if (block.dColumn != null && block.dColumn.startCellIndex == block.joinat)
					{
						// Merge Joined Table Columns.
						for (int j = 0; j < block.lineList.Count; j++)
						{
							TplLine line = block.lineList [j] ;
						
							if (!line.containsHGroup)
								continue ;

							Boolean hasMerged = false ;
							for (int k = 0; k < line.insertedRowList.Count; k++)
							{
								int rowIndex = line.insertedRowList [k] ;
								rowIndex = rowIndex - block.startRowIndex + blockRow ;

								Range leftRange = GetExcelRange (sheet, blocklastColum, rowIndex, 1, 1) ;
								if (leftRange.MergeArea != null)
									leftRange = GetExcelRange (sheet, leftRange.MergeArea.Column, leftRange.MergeArea.Row, 1,1) ;
								Range rightRange = GetExcelRange(sheet, blocklastColum + 1, rowIndex, 1, 1);
								if (rightRange.MergeArea != null)
									rightRange = GetExcelRange(sheet, rightRange.MergeArea.Column, rightRange.MergeArea.Row, 1, 1);
							
								if (leftRange.Text.Equals (rightRange.Text))
								{
									// Merge 
									GetExcelRange(sheet, leftRange.Column, 
									              leftRange.Row,
									              rightRange.Column + rightRange.Columns.Count - leftRange.Column, 
								               
									              Math.Min(rightRange.Rows.Count, leftRange.Rows.Count) 
										).Merge (true) ;
									
									hasMerged = true ;
								}
							}
						
							if (! hasMerged)
								break ;
						} // end for 
					} // end if 
					blocklastColum += block.colCount - block.joinat ;
					
				}
			}
		}
		*/
		public static void Clear(Spire.Xls.Worksheet sheet, int startRowIndex)
		{
			// sheet.get_Range ("A1", "B1").EntireColumn.Delete (XlDeleteShiftDirection.xlShiftToLeft) ;
			sheet.DeleteColumn (1);
			sheet.DeleteColumn (1);
			// sheet.DeleteColumn (1);
			// sheet.get_Range ("A1", "A1").EntireColumn.Delete (XlDeleteShiftDirection.xlShiftToLeft) ;
			
			// sheet.get_Range (RangeHelper.MakeCellName (1, 1), RangeHelper.MakeCellName (1, startRowIndex - 1)).EntireRow.Delete (Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp) ;
			
			
			
			for (int i = 1; i < startRowIndex - 1; i++)
			{
				sheet.DeleteRow (1);
			}
			
		}
/*
		public static void ClearByExcel(Worksheet sheet, int startRowIndex)
		{
			sheet.get_Range("A1", "B1").EntireColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
			// sheet.get_Range ("A1", "A1").EntireColumn.Delete (XlDeleteShiftDirection.xlShiftToLeft) ;

			sheet.get_Range(RangeHelper.MakeCellName(1, 1), RangeHelper.MakeCellName(1, startRowIndex - 1)).EntireRow.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

			/*
			for (int i = 1; i < startRowIndex; i++)
			{
				sheet.get_Range("A1", "A1").EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
			}
			*  /
		}
*/
/*		public static int MergeExcelRanges(Range currentCellRange, MergeOption mOption)
		{
			Worksheet sheet = currentCellRange.Worksheet;
			Range range = null;
			// clear value first.
			currentCellRange.Value2 = null;
			// currentCellRange.Text = "";

			switch (mOption)
			{
				case MergeOption.Up:

					// Select previous range.
					range = GetExcelRange(sheet, currentCellRange.Column, currentCellRange.Row - 1, 1, 1);

					if (range.MergeArea != null && range.MergeArea.Cells.Count > 1)
					{
						Range marea = range.MergeArea;
						int mWidth = currentCellRange.Column - marea.Column + 1;
						int mHeight = currentCellRange.Row - marea.Row + 1;


						if (mWidth < marea.Column)
							mWidth = marea.Column;

						if (mHeight < marea.Row)
							mHeight = marea.Row;

						range = GetExcelRange(sheet,
									  marea.Column, marea.Row,
									  mWidth,
									  mHeight);
						
						range.Merge(Missing.Value);
					}
					else
					{
						range = GetExcelRange(sheet,
									 currentCellRange.Column,
									 currentCellRange.Row - 1,
									 1, 2);
						range.Merge(Missing.Value);
					}

					return 1;
				case MergeOption.Left:
					// merge with Left cell.

					range = GetExcelRange(sheet, currentCellRange.Column - 1, currentCellRange.Row, 1, 1);

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
						
						range.Merge();
					}
					else
					{
						range = GetRange(sheet,
									 currentCellRange.Column - 1,
									 currentCellRange.Row,
									 2, 1);

						range.Merge();
					}

					// range.Merge ();
					return 1;
				case MergeOption.never:
				default:
					return 0;
			}
		}

	
		public static Range GetExcelRange(Worksheet sheet, int startColumn, int startRow, int columns, int rows)
		{
			
			return sheet.get_Range (RangeHelper.MakeCellName (startColumn, startRow),
									RangeHelper.MakeCellName(startColumn + columns - 1, startRow + rows - 1));
		}
*/	
        //add by chengzy
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="autoFit"></param>
        public void LoadReportNoOpen(string fileName, bool autoFit)
        {
			Workbook book = new Workbook();


			book.LoadFromFile(fileName);

			ReportBookTemplate tplbook = new ReportBookTemplate();
			tplbook.LoadTemplate(db, book, paramMap);
			tplbook.FillTemplate();

			List<ReportSheetTemplate> tplList = tplbook.sheetList;
			// tpl.Clear();

			// Open with Excel

			/*			xlapp.Visible = false;*/

			try
			{
				ClearReport(autoFit, book, tplList);


				// xlapp.Workbooks.Close ();
			}
			finally
			{
				book.Save();
				book.Dispose();

				RemoveWarning(fileName);


				/*				xlapp.DisplayAlerts = true ;
			
								xlapp.Visible = true;
				 */
//				Process.Start(fileName);
			}

        }
	}
}
