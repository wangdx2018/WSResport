using System ;
using System.Collections.Generic ;
// using Microsoft.Office.Interop.Excel ;
using Spire.Xls ;
using DataTable=System.Data.DataTable;
using Range=Spire.Xls.CellRange ;

namespace AFC.WorkStation.ExcelReport
{
	public class TplBlock
	{
		public int tplRowCount = 1 ;
		public int tplColumCount = 1 ;
		public string tableName ;
		public Range tplRange ;
		
		public int startRowIndex ;
		public int startColIndex ;
		public int rowCount ;
		public int colCount ;
		public bool copyOnly = false ;
		public int joinat = -1 ;

		public TplLine lastUsedLine = null ;
		public int lastUsedLineValueIndex = -1 ;
		
		public List<TplLine> lineList = new List<TplLine> () ;
		public GroupDataHolder holder = new GroupDataHolder ();
		public List<GroupValueSearchKey> gkeyList = new List<GroupValueSearchKey>(64*1024);
		public TplCloumn dColumn = null ;
		public Dictionary<GroupValueSearchKey, bool> countedMap = new Dictionary<GroupValueSearchKey, bool>(64*1024) ;
		public string name ;
		public bool updateAllRow = false ;
		public string tplColTableName = null; 
		public bool dCloumnsCreated = false ;
        //the default number of  empty rows 
        //public int defaultEmptyRowsCount = 0;
	    
	    //
        public string emptyTableName;
		public bool isChart = false;
		public string chartDataBlockName = null;
		public TplBlock chartDataBlock ;
		public bool chartSeriesFrom = true ;

		public void InitDColumn ()
		{
			dColumn = FindDColumn();
		}

		private TplCloumn FindDColumn ()
		{
			TplCloumn dcol = new TplCloumn ();
			dcol.gCols = 0 ;
			
			for (int i = 0; i < lineList.Count; i++)
			{
				TplLine line = lineList [i] ;
				
				// int startIndex = -1 ;
				for (int j = 0; j < line.cellList.Count; j++)
				{
					TplCell cell = line.cellList [j] ;
					
					if (cell.align == GroupAlign.hGroup)
					{
						// Only One DColumn 
						if (dcol.gCols <= 0)
						{
							dcol.startCellIndex = j ;
							dcol.startColIndex = cell.lastColIndex ;
							
						}
		
						dcol.gCols ++ ;
						
					}
					
				}
				
				
				if (dcol.gCols > 0)
				break ;
			}

			if (dcol.gCols <= 0)
				return null ;
			
			dcol.tplLastColCount = tplColumCount - dcol.startCellIndex - dcol.gCols ;
			if (dcol.tplLastColCount <= 0)
				dcol.tplLastColCount = 1 ;
			
			// find out gcells
			dcol.tplRange = RangeHelper.GetRange (tplRange.Worksheet, dcol.startColIndex
			                                      , tplRowCount, dcol.gCols, lineList.Count);

			for (int i = 0; i < lineList.Count; i++)
			{
				TplLine line = lineList [i] ;

				for (int j = 0; j < line.cellList.Count; j++)
				{
					TplCell cell = line.cellList [j] ;
					
					if (cell.align == GroupAlign.hGroup)
					{
						TplCell newCell = cell.Copy();
						dcol.cellList.Add(newCell);

						if (!dcol.groupColList.Contains(cell.tplGroupColName))
						{
							dcol.groupColList.Add(cell.tplGroupColName);
							dcol.groupColIndexList.Add (-1);
						}
						
					
					}
				}
				
			}
			
			return dcol ;
		}
		
		public int CreateDColumns (DataTable table)
		{
			if (dColumn == null)
				return 0 ;
			
			// Current colCount include gCols, so substract it.
			colCount -= dColumn.gCols ;
			
			// try to insert dColumns into Template. 
			// dColumn.CheckColumn ()
			if (table == null || table.Rows.Count <= 0)
				return 0 ;

			for (int i = 0; i < table.Rows.Count; i++)
			{
				dColumn.CheckEachColumn (this, holder, 0, table, i);
				// dColumn.InsertColumn (this, holder, table, i, false);
			}
			
			dCloumnsCreated = true ;
			
			// remove data from template 
			// Insert new Col in Template.
			Range tplColRange = RangeHelper.GetRange(tplRange.Worksheet,
													  dColumn.startColIndex + dColumn.gCols, tplRange.Row,
													  tplColumCount - dColumn.startCellIndex - dColumn.gCols - 1, 
			                                         // 50,
			                                         tplRowCount);

			tplColRange.Copy (
				RangeHelper.GetRange (tplRange.Worksheet,
									   dColumn.startColIndex, tplRange.Row,
				                      tplColumCount - dColumn.startCellIndex - dColumn.gCols - 1,
				                      // 50,
				                      tplRowCount), true, true) ;
			/* 
			RangeHelper.InsertCopyRange(tplRange.Worksheet, tplColRange,
										tplColumCount - dColumn.startCellIndex - dColumn.gCols, 
			                            tplRowCount,
										startColIndex , tplColRange.Row,
										XlInsertShiftDirection.xlShiftToRight, tplColumCount);
			*/
			tplColumCount -= dColumn.gCols ;
			for (int i = 0; i < lineList.Count; i++)
			{
				TplLine line = lineList [i] ;
				
				line.cellList.RemoveRange (dColumn.startCellIndex, dColumn.gCols);
				line.tplCellCount -= dColumn.gCols ;
				line.tplRange = RangeHelper.GetRange (tplRange.Worksheet,
				                                      3, line.tplRange.Row, 
				                                      tplColumCount, 1) ;
			}
			// Refresh Line.TplRange ;
			// RefreshLineTplRanges(block, 1);
		
			return 1 ;
			
		}
		
		public int FillBlock (DataTable table)
		{
			int rowIndex = startRowIndex ;
			for (int valueIndex = 0; valueIndex < table.Rows.Count; valueIndex++)
			{
				countedMap.Clear (); 
				// fill key for each line, and count Data.
				// Check is Need Row ;
				if (dColumn != null && ! dCloumnsCreated)
				{
					dColumn.CheckColumn (this, holder, rowIndex,
					                     table, valueIndex) ;
				}
				
				// fill key for each line, and count Data.
				for (int j = 0; j < gkeyList.Count; j++)
				{
					GroupValueSearchKey gkey = gkeyList [j] ;
					
					if (gkey.rKey == null)
					{
						SearchKey key = new SearchKey();
						key.colName = gkey.valueColName;
						gkey.rKey = ReusedKey.FindReusedKey(key);
					}
					
					gkey = gkey.Copy () ;
					if (gkey.key != null)
						gkey.key.FillKey (table, valueIndex);
					
					holder.AddValue (countedMap, gkey, table, valueIndex);
				}
				
				// fill Data 
				for (int j = 0; j < lineList.Count; j++)
				{
					TplLine line = lineList[j];
					/*
					if (j == 4)
						j = 4 ;
					*/
					int nl = line.FillLine(holder, rowIndex, table, valueIndex);
					if (nl > 0)
					{
						FillLastLine(j, lastUsedLine, rowIndex - 1, table, lastUsedLineValueIndex);

						lastUsedLine = line ;
						lastUsedLineValueIndex = valueIndex ;

						// is LastLine, update current line
						if (valueIndex + 1 >= table.Rows.Count)
						{
							FillLastLine(j, line, rowIndex, table, lastUsedLineValueIndex);
						}
					
					}
					
				
					/*else
					{
						// Ensure Each Column is OK.
						if (dColumn != null && lastUsedLine != null && lastUsedLine == line)
						{
							if (! lastUsedLine.containsHGroup)
								lastUsedLine.UpdateRowData (holder, rowIndex - 1, table, valueIndex);
							
							if (updateAllRow)
							{
								for (int k = 0 ; k <= j ; k ++)
								{
									TplLine hLine = lineList [k] ;
									if (hLine.containsHGroup && hLine.insertedRowList.Count > 0)
									{
										hLine.UpdateRowData(holder, 
										                    hLine.insertedRowList[hLine.insertedRowList.Count - 1], 
										                    table, valueIndex);
									}
								}
							}
							
						}
					}*/
					rowIndex += nl ;
					rowCount += nl ;
				}
				
			}

			MergeHGroupCells () ;

            //if the table is custum empty ,ignor the  MergeVGroupCells
            if (table.ExtendedProperties.ContainsKey("TableType")
                && table.ExtendedProperties["TableType"].ToString () == "CustumEmpty")
            {
                return rowCount;
            }  
		    
		    MergeVGroupCells () ;
			return rowCount;
		}

		private void FillLastLine (int currentLineIndex, TplLine updateLine, int rowIndex, DataTable table, int valueIndex)
		{
			if (/*dColumn != null && */updateLine != null) //&& lastUsedLine == line)
			{
				//		if (lastUsedLine.containsHGroup)
				updateLine.UpdateRowData(holder, rowIndex, table, valueIndex);

				if (updateAllRow)
				{
					for (int k = 0; k <= currentLineIndex; k++)
					{
						TplLine hLine = lineList[k];
						if (hLine.containsHGroup && hLine.insertedRowList.Count > 0)
						{
							hLine.UpdateRowData(holder,
							                    hLine.insertedRowList[hLine.insertedRowList.Count - 1],
							                    table, valueIndex);
						}
					}
				}

			}
		}

		public void MergeHGroupCells ()
		{
			for (int i = 0; i < lineList.Count; i++)
			{
				TplLine line = lineList [i] ;
				if (! line.containsHGroup)
					continue ;

				for (int j = 0; j < line.insertedRowList.Count; j++)
				{
					int rowIndex = line.insertedRowList [j] ;

					object lastValue = null ;
					int colIndex = dColumn.startColIndex ;
					// if (line.cellList [colIndex].mOption != MergeOption.Left)
					//	continue ;
					// ingore this line ;
				
					for (int k = 0 ; k < dColumn.insertCount ; k ++)
					{
						Range cell = RangeHelper.GetCell (tplRange.Worksheet, colIndex , rowIndex) ;
						object value = cell.Value2 ;
						
						
						if (lastValue != null && 
						    Equals (lastValue, value))
						{
							/* remove after debug.
							if (colIndex == 27)
								Console.WriteLine ("colINdex=27");
							*/
							// clear
							// judge if last row is last hgrouoped row.
							if (i == lineList.Count - 1 || lineList [i + 1].containsHGroup)
							{
								RangeHelper.MergeRanges (cell, MergeOption.Left) ;
							}
							else
							{
								// check is this column first column in hgrouped columns.
								if (k % dColumn.gCols > 0)
									RangeHelper.MergeRanges(cell, MergeOption.Left);
							}
						}
							
						lastValue = value ;
						colIndex ++ ;
					}

					// repair unmerged range 
					int afterIndex = dColumn.startCellIndex + dColumn.insertCount;
					for (int k = afterIndex ; 
						k < line.cellList.Count ; k ++)
					{
						TplCell cell = line.cellList [k] ;
						
						if (cell.mOption == MergeOption.Left)
						{
							Range range = RangeHelper.GetCell (tplRange.Worksheet, colIndex, rowIndex) ;
							RangeHelper.MergeRanges (range, MergeOption.Left) ;
						}
						colIndex += cell.acrossColumns ;
					}
				}
				
				
				
			}
		}

		public void MergeVGroupCells()
		{
			for (int i = 0; i < lineList.Count; i++)
			{
				TplLine line = lineList[i];
				/* 
				if (!line.containsHGroup)
					continue;
				*/
				int colIndex = line.colIndex ;
				for (int j = 0; j < line.cellList.Count; j++)
				{
					TplCell cell = line.cellList [j] ;
					if (cell.mOption != MergeOption.Up || 
					    (cell.align != GroupAlign.none && 
					     cell.align != GroupAlign.always))
					{
						colIndex += cell.acrossColumns ;
						continue ;
					}
					for (int k = 0; k < line.insertedRowList.Count; k++)
					{
						int rowIndex = line.insertedRowList[k];

						CellRange range = RangeHelper.GetCell (tplRange.Worksheet, colIndex, rowIndex) ;
						
						if (! range.HasMerged)
							RangeHelper.MergeRanges (range, MergeOption.Up) ;
					}

					colIndex += cell.acrossColumns;
				}
				
			}
		}

		public void SetupChart ()
		{
			if (!isChart || chartDataBlock == null)
				return ;

			Worksheet sheet = chartDataBlock.tplRange.Worksheet ;

			Chart chart = tplRange.Worksheet.Charts [0] ;
			// bool series = chart.SeriesDataFromRange ;
			if (chartDataBlock.rowCount <= 0)
			{

				chart.Series.Clear ();
				chart.DataRange = RangeHelper.GetRange(sheet, chart.LeftColumn, chart.TopRow, 1, 1);
			}
			else
			{
				chart.DataRange = RangeHelper.GetRange (sheet,
				                                        chartDataBlock.startColIndex + 1, chartDataBlock.startRowIndex,
				                                        chartDataBlock.colCount - 1, chartDataBlock.rowCount) ;
			}

			chart.SeriesDataFromRange = chartSeriesFrom;
			int moveRows = startRowIndex - chart.TopRow ;
			
			
			chart.TopRow += moveRows;
			chart.BottomRow += moveRows ;
			
			// chart.AutoSize = true ;
			// chart.AutoScaling = true ;
		
		}
	}

	public class TplCloumn
	{
		public int nextCloIndex ;
		public int tplLastColCount = 1 ;
		public int insertCount = 0 ;
		public int startCellIndex = 0 ;
		public int gCols = 1 ;
		public Range tplRange ;
		public List<TplCell> cellList = new List<TplCell> () ;
		public int startColIndex ;
		public List<string> groupColList = new List<string> ();
		internal List<int> groupColIndexList = new List<int> ();
		private DataTable lastUsedTable = null ;
		public List<object []>lastValueMap = new List<object[]> ();
		
		
		public ReportSheetTemplate tpl ;
		
		public int CheckColumn (TplBlock block, GroupDataHolder holder, int currentRowIndex, System.Data.DataTable table, int valueIndex)
		{
			// if grouped Value changed.
			// if value has be inserted.
			// Insert Cell At EachLine
			// Update TplLine Info.
			int startIndex = IsNeedNewCol (currentRowIndex, holder, table, valueIndex) ;
			
			if (startIndex < 0)
				return 0 ;
			// Console.WriteLine("---- Start of " + valueIndex);

			
			InsertColumn (block, holder, table, valueIndex, true) ;

			return 1 ;
		}

		public void CheckEachColumn(TplBlock block, GroupDataHolder holder, int currentRowIndex, System.Data.DataTable table, int valueIndex)
		{
			for (int i = 0 ; i < gCols ; i ++)
			{
				bool colInserted = false ;
				// Check each cell in this columns 
				for (int j = 0; j < block.lineList.Count; j++)
				{
					
					TplLine line = block.lineList [j] ;
					TplCell cell = line.cellList [startCellIndex + i] ;
					if (cell.align != GroupAlign.hGroup)
						continue ;

					if (! colInserted)
					{
						bool needNew = cell.IsNeedNewCell(holder, cell.hgOption, 0, currentRowIndex, table, valueIndex);
					
						if (needNew)
						{
							InsertOneColumn (block, i, holder, table, valueIndex, false);
							colInserted = true ;
						}
					}
					else
					{
						if ((cell.hgOption & InsertOption.BeforeChange) != 0)
							continue ;
						
						// set last grouped value 
						cell.lastGroupedValue = cell.GetGroupValue (holder, table, valueIndex) ;
					}
				}
			}
		}
		public void InsertOneColumn (TplBlock block, int colIndex, GroupDataHolder holder, DataTable table, int valueIndex, bool hasData)
		{
			if (hasData)
			{
				// block.startRowIndex ;
				Range colRange = RangeHelper.GetRange(tplRange.Worksheet,
													   startColIndex + colIndex, block.startRowIndex,
													   1, block.rowCount);
				// Insert new ;
				RangeHelper.InsertCopyRange(tplRange.Worksheet, colRange,
											 1, block.rowCount,
											 startColIndex + gCols + insertCount, block.startRowIndex,
											 XlInsertShiftDirection.xlShiftToRight, tplLastColCount);
			}
			// Insert new Col in Template.
			Range tplColRange = RangeHelper.GetRange(tplRange.Worksheet,
													  startColIndex + colIndex, block.tplRange.Row,
													  1, block.tplRowCount);

			RangeHelper.InsertCopyRange(tplRange.Worksheet, tplColRange,
										1, block.tplRowCount,
										startColIndex + gCols + insertCount, tplColRange.Row,
										XlInsertShiftDirection.xlShiftToRight, tplLastColCount);
			// Refresh Line.TplRange ;
			RefreshLineTplRanges(block, 1);

			block.tplColumCount += 1;
			block.colCount += 1;

			// Insert cell into exsit lineList.
			for (int lineIndex = 0; lineIndex < block.lineList.Count; lineIndex++)
			{
				TplLine line = block.lineList[lineIndex];

				int cellIndex = startCellIndex + colIndex;

				TplCell cell0 = line.cellList[cellIndex];

				TplCell cell = cell0.Copy();
				cell.lastColIndex += 1;

				
				line.cellList.Insert(startCellIndex + gCols + insertCount, cell);

				/* 
				 if (cell.useExcelFormula)
				{
					cell.tplRange = cell0.tplRange ;
				}
				*/
				if (cell.formula != null)
				{
					for (int keyIndex = 0; keyIndex < cell.formula.keyList.Count; keyIndex++)
					{
						GroupValueSearchKey gkey = cell.formula.keyList[keyIndex];
						if (gkey.rKey == null)
						{
							SearchKey key0 = new SearchKey ();
							key0.colName = gkey.valueColName ;
							
							gkey.rKey = ReusedKey.FindReusedKey (key0);
							
						}
						SearchKey key = gkey.key;
						while (key != null)
						{
							if (IsGroupedColumn(key.colName))
							{
								key.keyValue = RangeHelper.GetColValue(table, valueIndex, key.colName);
								key.isFixedValue = true;
							}
							key = key.nextKey;

						}

						block.gkeyList.Add(gkey.Copy());
						if (gkey.key != null)
							gkey.key.FillKey(table, valueIndex);

						block.holder.AddValue(block.countedMap, gkey, table, valueIndex);
					}
				}
				/* 
				else if (cell.hgOption != InsertOption.never)
				{
					// set fixed text 
					cell.tplTextContent = Convert.ToString(RangeHelper.GetColValue (table, valueIndex, cell.tplValueColName)) ;
				}
				*/
				
				cell.align = GroupAlign.none;

				Console.WriteLine("Inserted hg Line[" + lineIndex + "]cell[" + cellIndex + "] = " + cell.formula);

				/* update Row Value */
				if (lineIndex < block.rowCount)
				{
					Range cellRange = RangeHelper.GetCell(tplRange.Worksheet, startColIndex + gCols + insertCount,
														   block.startRowIndex + lineIndex);

					cell.WriteCell(tpl, holder, cellRange, table, valueIndex);
				}


			}
			// Console.WriteLine ("---- End of " + valueIndex);
			// increment next 

			nextCloIndex += 1;
			insertCount++;
			
		}
		
		public void InsertColumn (TplBlock block, GroupDataHolder holder, DataTable table, int valueIndex, bool hasData)
		{
			// do insert 
			if (insertCount > 0)
			{
				if (hasData)
				{
					// block.startRowIndex ;
					Range colRange = RangeHelper.GetRange (tplRange.Worksheet,
					                                       startColIndex + insertCount - gCols , block.startRowIndex,
					                                       gCols, block.rowCount);
					// Insert new ;
					RangeHelper.InsertCopyRange (tplRange.Worksheet, colRange,
					                             gCols, block.rowCount,
					                             startColIndex + insertCount, block.startRowIndex, 
					                             XlInsertShiftDirection.xlShiftToRight, tplLastColCount) ;
				}
				// Insert new Col in Template.
				Range tplColRange = RangeHelper.GetRange (tplRange.Worksheet,
				                                          startColIndex + insertCount - gCols, block.tplRange.Row,
				                                          gCols, block.tplRowCount) ;

				RangeHelper.InsertCopyRange(tplRange.Worksheet, tplColRange, 
				                            gCols, block.tplRowCount,
				                            startColIndex + insertCount, tplColRange.Row, 
				                            XlInsertShiftDirection.xlShiftToRight, tplLastColCount);
				// Refresh Line.TplRange ;
				RefreshLineTplRanges (block, gCols) ;
				
				block.tplColumCount += gCols ;
				block.colCount += gCols ;
			}
			
			
			// Insert cell into exsit lineList.
			for (int lineIndex = 0; lineIndex < block.lineList.Count; lineIndex++)
			{
				TplLine line = block.lineList [lineIndex] ;

				for (int j = 0 ; j < gCols; j++)
				{
					int cellIndex = startCellIndex + (insertCount > 0 ? (insertCount - gCols) : 0) + j;
					
					TplCell cell = line.cellList [cellIndex] ;
					
					/* if (cell.lastColIndex != nextCloIndex - (insertCount > 0 ? 1 : 0)) */
					// if (cell.lastColIndex < nextCloIndex || cell.lastColIndex >= nextCloIndex + gCols)
					// 	continue ;
				
					//	if (lineIndex == 2)
					//		lineIndex = 2 ;
					
					if (insertCount > 0)
					{
						cell = cell.Copy () ;
						cell.lastColIndex += gCols ;
				
						line.cellList.Insert (cellIndex + gCols, cell);
					}
					
					if (cell.formula != null)
					{
						for (int keyIndex = 0; keyIndex < cell.formula.keyList.Count; keyIndex++)
						{
							GroupValueSearchKey gkey = cell.formula.keyList[keyIndex];
							SearchKey key = gkey.key;
							while (key != null)
							{
								if (IsGroupedColumn(key.colName))
								{
									key.keyValue = RangeHelper.GetColValue(table, valueIndex, key.colName);
									key.isFixedValue = true;
								}
								key = key.nextKey;

							}

							block.gkeyList.Add(gkey.Copy ());
							if (gkey.key != null)
								gkey.key.FillKey (table, valueIndex);
							
							block.holder.AddValue(block.countedMap, gkey, table, valueIndex);
						}
					}
					else
						if (cell.hgOption != InsertOption.never)
					{
						cell.tplTextContent = Convert.ToString(cell.GetValueByIndex (valueIndex, table)) ;
					}
					
					cell.align = GroupAlign.none;
				
					Console.WriteLine ("Inserted hg Line[" + lineIndex + "]cell[" + cellIndex + "] = " + cell.formula);
				
					/* update Row Value */
					if (lineIndex < block.rowCount)
					{
						Range cellRange = RangeHelper.GetCell (tplRange.Worksheet, startColIndex + (insertCount/* == 0 ? 0 : insertCount - 1*/) + j,
						                                       block.startRowIndex + lineIndex) ;
					
						cell.WriteCell (tpl, holder, cellRange, table, valueIndex);
					}
						
				}
			
			}
			// Console.WriteLine ("---- End of " + valueIndex);
			// increment next 
			
			nextCloIndex += gCols ;
			insertCount += gCols ;
		}

		private void RefreshLineTplRanges (TplBlock block, int colCount)
		{
			for (int i = 0; i < block.lineList.Count; i++)
			{
				TplLine line = block.lineList [i] ;
				
				line.tplRange = RangeHelper.GetRange (line.tplRange.Worksheet,
				                                      line.tplRange.Column,
				                                      line.tplRange.Row,
				                                      line.tplCellCount + insertCount,
													  colCount
					) ;
				line.tplCellCount += colCount;
			}
		}

		private Dictionary<string, bool> groupedColMap = new Dictionary<string, bool> (64); 
		private bool IsGroupedColumn (string colName)
		{
			bool isGCol ;
			if (groupedColMap.TryGetValue (colName, out isGCol))
				return isGCol ;
			
			for (int i = 0; i < cellList.Count; i++)
			{
				TplCell cell = cellList [i] ;
				if (cell.align == GroupAlign.hGroup && cell.tplGroupColName == colName)
				{
					groupedColMap.Add (colName, true);
					return true ;
				}
			}
			groupedColMap.Add(colName, false);
			return false ;
		}

		private int IsNeedNewCol (int currentRowIndex, GroupDataHolder holder, System.Data.DataTable table, int valueIndex)
		{
			
			object [] vList = new object[groupColList.Count];
			if (lastUsedTable != table)
			{
				lastUsedTable = table ;
				// Clear to -1 ;
				for (int i = 0; i < groupColIndexList.Count; i++)
				{
					groupColIndexList [i] = table.Columns.IndexOf (groupColList [i]) ;
				}
			}
			for (int i = 0; i < groupColList.Count; i++)
			{
				/* string colName = groupColList [i] ;  */

				vList[i] = RangeHelper.GetColValue(table, valueIndex, groupColIndexList[i]);
			}

			for (int i = 0; i < lastValueMap.Count; i++)
			{
				object[] lvList = lastValueMap [i] ;
				
				if (CompareArray (vList, lvList))
					return -1 ;
			}
			
			// Save new value list 
			lastValueMap.Add (vList);
			
			/* for (int i = 0; i < cellList.Count; i++)
			{
				TplCell cell = cellList[i];
				if (cell.align != GroupAlign.hGroup)
					continue ;
				if (! cell.IsNeedNewCell(holder, InsertOption.afterChange, cell.lastColIndex, currentRowIndex, table, valueIndex))
					index = i ;
				
			}
			*/
			return 0 ;
			
		}
		
		private static bool CompareArray (object [] a1, object [] a2)
		{
			if (a1.Length != a2.Length)
				return false ;

			for (int i = 0; i < a1.Length; i++)
			{
				if (! Equals (a1 [i], a2 [i]))
					return false ;
			}
			return true ;
		}
	}

	public class TplLine
	{
		public int tplCellCount = 0;
		public Range tplRange;
		public List<TplCell> cellList = new List<TplCell>();
		public int colIndex = 1;
		public int rowIndex = 1;
		public int rowCount = 0;
		public InsertOption iOption = InsertOption.afterChange;
		public bool containsHGroup = false ;
		
		public List<int> insertedRowList = new List<int> (16*1024);
		public ReportSheetTemplate tpl ;
		public int FillLine(GroupDataHolder holder, int currentRowIndex, System.Data.DataTable table, int valueRowIndex)
		{
			// Judge Dose Need New Line 

			int startIndex = IsNeedNewLine(holder, currentRowIndex, table, valueRowIndex);

			if (startIndex < 0)
				return 0;
			if (startIndex >= 0)
			{
				// Insert New Line 
				/* Range newLineRange = */
				RangeHelper.InsertCopyRange(tplRange.Worksheet, tplRange, cellList.Count, 1, colIndex, currentRowIndex, XlInsertShiftDirection.xlShiftDown);

				insertedRowList.Add(currentRowIndex);
			}


			UpdateLine (currentRowIndex, holder, startIndex, table, valueRowIndex, MergeOption.Left, false) ;
			UpdateLine (currentRowIndex, holder, startIndex, table, valueRowIndex, MergeOption.Up, true) ;

			return 1;
		}

		private void UpdateLine (int currentRowIndex, GroupDataHolder holder, int startIndex, DataTable table, int valueRowIndex, MergeOption mo, bool updateMergeOnly)
		{
			int cellColIndex = cellList[0].lastColIndex;
			for (int i = 0; i < cellList.Count; i++)
			{
				TplCell cell = cellList[i];
				
				int merge = 0;

				if (cell.mOption == mo && 
				    (// merge before.
				    i < startIndex ||
				    // merge after. next line is new line
				    (i >= startIndex && (iOption & InsertOption.BeforeChange) != 0) || 
				    // ignore group align option.
				    cell.align == GroupAlign.none))
					// do Merge 
				{
					Range cellRange = RangeHelper.GetCell(tplRange.Worksheet, cellColIndex, currentRowIndex);

					merge = cell.DoMerge(currentRowIndex, cellRange);
				}

				if (merge == 0 && ! updateMergeOnly)
				{
					if (cell.formula == null)
					{
						Range cellRange = RangeHelper.GetCell(tplRange.Worksheet, cellColIndex, currentRowIndex);

						cell.WriteCell(tpl, holder, cellRange, table, valueRowIndex);
					}
				}
				/*
				else
					// todo: Remove this noused line?
					cell.lastGroupedValue = cell.GetValue(holder, table, valueRowIndex);
				*/
				cellColIndex += cell.acrossColumns;
			}
		}


		public void UpdateRowData(GroupDataHolder holder, int currentRowIndex, System.Data.DataTable table, int valueRowIndex)
		{

			int cellColIndex = cellList[0].lastColIndex;
			for (int i = 0; i < cellList.Count; i++)
			{
				TplCell cell = cellList[i];
				if (cell.formula != null)
				{
					Range cellRange = RangeHelper.GetCell(tplRange.Worksheet, cellColIndex, currentRowIndex);
				
					cell.WriteCell(tpl, holder, cellRange, table, valueRowIndex);
				}
				cellColIndex += cell.acrossColumns;
			}

		}
		public int IsNeedNewLine(GroupDataHolder holder, int currentRowIndex, System.Data.DataTable table, int valueRowIndex)
		{
			if (iOption == InsertOption.never)
				return -1;

			if (valueRowIndex == 0 && (iOption & InsertOption.onfirst) != 0)
			{
				return 0;
			}

			if (valueRowIndex == table.Rows.Count - 1 && (iOption & InsertOption.onLast) != 0)
			{
				return 0;
			}

			if ((iOption & InsertOption.afterChange) == 0 &&
			    (iOption & InsertOption.BeforeChange) == 0 &&
			    (iOption & InsertOption.always) == 0)
				return -1;

			for (int i = 0; i < cellList.Count; i++)
			{
				TplCell cell = cellList[i];
				if (cell.align != GroupAlign.hGroup && cell.IsNeedNewCell(holder, iOption, cell.lastColIndex, currentRowIndex, table, valueRowIndex))
				{
					return i;
				}
			}
			return -1;
		}
	}

	public class TplCell
	{
	
		/* Range of Template */
		public Range tplRange = null;

		/* Grouped action align */
		public GroupAlign align = GroupAlign.none;
		/* A cell can across more then 1 column. those row will be skiped.*/
		public int acrossColumns = 1;
		/* Merge after new line insert, or before next new line insert. */
		public MergeOption mOption = MergeOption.never;

		/* mapped colName */
		public string tplGroupColName = string.Empty;
		public string tplValueColName = string.Empty;
		/* mapped value format string */
		public string tplFormat = null;
		/* original text content */
		public string tplTextContent = string.Empty;

		// runtime info.
		// How many rows/columns merged(or need to merge).
		public int gMergeCount;
		// the column index of the cell last inserted.
		public int lastColIndex;
		// the row index of the cell last inserted.
		public int lastRowIndex;
		// the value(from data table) of the cell last inserted.
		public object lastGroupedValue;

	    //default value
        public string tplDefaultContent = string.Empty;
	    
		public Range lastCellRange = null;

		public CellForumla formula = null;

		public Dictionary<object, bool> lastValueMap = new Dictionary<object, bool>();

		public InsertOption hgOption = InsertOption.never;
		public bool useR1C1Formula = false ;
		private ReusedKey rGroupKey ;
		
		private DataTable lastUsedTable = null ; 
		private int lastUsedColIndex = -1 ;
		public object GetValue(GroupDataHolder holder, DataTable table, int rowIndex)
		{
			if (formula != null)
				return formula.GetValue(holder, table, rowIndex);

			if (hgOption != InsertOption.never) //  && align == GroupAlign.hGroup)
			{
				// if this cell is hg col and not use formula, use text content as value;
				return tplTextContent ;
			}

			object value = GetValueByIndex (rowIndex, table) ;

			if (value == null)
				return tplTextContent;
			else
				return value;
		}

		internal object GetValueByIndex (int rowIndex, DataTable table)
		{
			if (lastUsedTable != table)
			{
				lastUsedColIndex = table.Columns.IndexOf (tplValueColName) ;
				lastUsedTable = table ;
			}

			return RangeHelper.GetColValue(table, rowIndex, lastUsedColIndex) ;
		}

		public bool IsNeedNewCell(GroupDataHolder holder, InsertOption nlOption, int currentColIndex, int currentRowIndex, System.Data.DataTable table,
		                          int valueRowIndex)
		{
			if (align == GroupAlign.always)
				return true ;
			
			if (align == GroupAlign.none)
				return (nlOption & InsertOption.always) != 0 ;
			
			object value = GetGroupValue (holder, table, valueRowIndex) ;
			switch (align)
			{
				case GroupAlign.vGroup:
				case GroupAlign.hGroup:
					if ((nlOption & InsertOption.BeforeChange) != 0)
					{
						if (valueRowIndex + 1 >= table.Rows.Count)
							// last line 
							return true;

						object value1 = GetGroupValue(holder, table, valueRowIndex);

						object value2 = GetGroupValue(holder, table, valueRowIndex + 1);
						return !Equals(value1, value2);
					}
					else
					{
						
						if (lastGroupedValue == null || ! lastGroupedValue.Equals (value))
						{
							lastGroupedValue = value ;
							return true ;
						}
						else
							return false ;
					}

				/*case GroupAlign.hGroup:
					// todo: check hGroup Cells. dose need a New Row.
					if (value == null)
						return true;
					bool flag;
					if (lastValueMap.TryGetValue(value, out flag))
						return false;
					else
						lastValueMap.Add(value, true);
					return true;*/
				default :
					break ;
			}
			return false;
		}

		public object GetGroupValue (GroupDataHolder holder, DataTable table, int valueIndex)
		{
			if (rGroupKey == null)
			{
				SearchKey key = new SearchKey();
				key.colName = this.tplGroupColName;
				rGroupKey = ReusedKey.FindReusedKey (key) ;
			}
			return rGroupKey.GetReusedValue(table, valueIndex);
			// return RangeHelper.GetColValue(table, valueIndex, tplGroupColName); 
		}

		public int DoMerge(int currentRowIndex, Range currentCellRange)
		{
			if (currentCellRange == null)
				return 0;
			lastCellRange = currentCellRange;
			
			return RangeHelper.MergeRanges (currentCellRange, mOption) ;
		}

		public void WriteCell(ReportSheetTemplate tpl, GroupDataHolder holder, CellRange currentCellRange, DataTable table, int valueRowIndex)
		{
			if (useR1C1Formula)
			{
				currentCellRange.FormulaR1C1 = "=" + tplTextContent;
				return ; // do noting 
			}
			object value = GetValue(holder, table, valueRowIndex);

            if (value == null || value == string.Empty || value == System.DBNull.Value )
                value = tplDefaultContent;
		    
			// update value. 
			RangeHelper.UpdateCellValue (tpl, currentCellRange, value, tplFormat);
			
			
			if (align == GroupAlign.vGroup)
				lastGroupedValue = GetGroupValue (holder, table, valueRowIndex) ;
			
			lastCellRange = currentCellRange;
		}

	

		public TplCell Copy()
		{
			TplCell cell = new TplCell();
			cell.align = align;
			cell.acrossColumns = acrossColumns;
			cell.lastColIndex = lastColIndex;
			cell.lastRowIndex = lastRowIndex;
			cell.tplGroupColName = tplGroupColName;
			cell.tplValueColName = tplValueColName;
			cell.tplFormat = tplFormat;
			cell.tplTextContent = tplTextContent;
            cell.tplDefaultContent = tplDefaultContent;		    
		    cell.useR1C1Formula = useR1C1Formula ;
			cell.hgOption = hgOption ;
			if (formula != null)
				cell.formula = formula.Copy();
			cell.lastGroupedValue = null;
			return cell;
		}
	}
}
