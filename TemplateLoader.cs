using System ;
using System.Collections.Generic ;
// using Microsoft.Office.Interop.Excel ;
using System.Drawing ;
using Spire.Xls ;
using Spire.Xls.Collections ;
using Range=Spire.Xls.CellRange ;
namespace AFC.WorkStation.ExcelReport
{
	public class TplLoader
	{
		private static int MAX_ROW_COUNT = 300 ;

		public static ReportSheetTemplate ParseTemplate(Worksheet sheet, int sheetIndex)
		{
			
			/* Range topRange = RangeHelper.GetRange (sheet, 1, 1, 1, 1).EntireColumn ; */
			ReportSheetTemplate tpl = new ReportSheetTemplate () ;
			tpl.sheet = sheet ;
			
			// Load pics
			PicturesCollection pictures = sheet.Pictures ;

			List<Rectangle> pics = new List<Rectangle>();

			for (int i = 0; i < pictures.Count; i++)
			{
				ExcelPicture picture = pictures [i] ;
				pics.Add(new Rectangle(picture.Left, picture.Top, picture.Width, picture.Height));
			}
			
			tpl.pics = pics ;
			// clear pictures.
			sheet.Pictures.Clear ();
						
			int lastValuedRowIndex = 1 ;
			
			for (int rowIndex = 1; rowIndex < MAX_ROW_COUNT; rowIndex++)
			{
				Range range = RangeHelper.GetRange (sheet, 1, rowIndex, 1, 1) ;
				string text = (string) range.Value2 ;
				if (string.IsNullOrEmpty (text))
					continue ;

				lastValuedRowIndex = rowIndex ;
				
				if (text.Equals ("sql", StringComparison.CurrentCultureIgnoreCase))
				{
					// here is a SQL.
					range = RangeHelper.GetRange (sheet, 2, rowIndex, 1, 1) ;
					string tableName = range.Value2 as string ;
					
					if (string.IsNullOrEmpty (tableName))
						continue ;

					range = RangeHelper.GetRange(sheet, 3, rowIndex, 1, 1);
					string sql = range.Value2 as string;
					
					if (string.IsNullOrEmpty (sql))
						continue ;
					
					// add sheet prefix to tableName
					if (tableName.IndexOf ('.') < 0)
						tableName = "S" + sheetIndex + "." + tableName ;
					tpl.sqlList [tableName] = sql ;
					
					
					continue ;
				}
				
				Dictionary<string, string> blockParams = ParseKeyValuePair (text) ;

				TplBlock block = new TplBlock () ;

				block.startColIndex = 2 ;
				block.startRowIndex = rowIndex ;
				block.colCount = block.tplColumCount = int.Parse (blockParams ["cols"]) ;
				if (blockParams.ContainsKey ("name"))
					block.name = blockParams ["name"] ;
				else
					block.name = "S" + sheetIndex + ".block" + (tpl.blockList.Count + 1) ;

				// parse chart params
				if (blockParams.ContainsKey ("ischart") &&
					"true".Equals(blockParams ["ischart"]))
				{
					block.isChart = true ;
					// parse dataBlock 
					if (blockParams.ContainsKey ("datablock"))
					{
						block.chartDataBlockName = blockParams["datablock"];

						// add sheet prefix to tableName
						if (block.chartDataBlockName.IndexOf('.') < 0)
							block.chartDataBlockName = "S" + sheetIndex + "." + 
								block.chartDataBlockName;
					}
					else
					{
						block.chartDataBlockName = "S" + sheetIndex + ".block2";
					}
					// find chartDataBlock 
					for (int i = 0; i < tpl.blockList.Count; i++)
					{
						TplBlock blk = tpl.blockList [i] ;
						if (blk.name.Equals (block.chartDataBlockName))
							block.chartDataBlock = blk ;
					}
						
					
					
					if (blockParams.ContainsKey("seriesfrom") && 
						"col".Equals (blockParams ["seriesfrom"]))
					{
						block.chartSeriesFrom = false ;
					}
				}
				int blockRows = block.tplRowCount = int.Parse (blockParams ["rows"]) ;

				lastValuedRowIndex += blockRows ;
				if (blockParams.ContainsKey ("copy"))
					block.copyOnly = "true".Equals (blockParams ["copy"]) ;

				
				block.tableName = blockParams ["table"] ;

				// add sheet prefix to tableName
				if (block.tableName.IndexOf('.') < 0)
					block.tableName = "S" + sheetIndex + "." + block.tableName;
				
				
				if (blockParams.ContainsKey("updateallrow"))
					block.updateAllRow = "true".Equals(blockParams["updateallrow"]);

				if (blockParams.ContainsKey("autofit") && blockParams["autofit"] == "true")
				{
					tpl.autoFit = true ;
				}

				if (blockParams.ContainsKey("joinat"))
				{
					if (! int.TryParse (blockParams ["joinat"], out block.joinat))
						block.joinat = -1 ;
				}

                //if (blockParams.ContainsKey("emptycount"))
                //{
                //    if (!int.TryParse(blockParams["emptycount"], out block.defaultEmptyRowsCount))
                //        block.defaultEmptyRowsCount  = 0;
                //}
                if (blockParams.ContainsKey("emptytable"))
                {
                	block.emptyTableName = blockParams["emptytable"];
					// add sheet prefix to tableName
					if (block.emptyTableName.IndexOf('.') < 0)
						block.emptyTableName = "S" + sheetIndex + "." + block.emptyTableName;
				
                }
				if (blockParams.ContainsKey("coltable"))
				{
					block.tplColTableName = blockParams["coltable"];
					// add sheet prefix to tableName
					if (block.tplColTableName.IndexOf('.') < 0)
						block.tplColTableName = "S" + sheetIndex + "." + block.tplColTableName;
				}
				
				block.tplRange = RangeHelper.GetRange (sheet, block.startColIndex, block.startRowIndex,
				                                       block.colCount, blockRows) ;

				if (block.copyOnly)
					// Just return directly.
				{
					tpl.blockList.Add (block) ;
					continue ;
				}


				for (int i = 0; i < blockRows; i++)
				{
					TplLine line = ParseLine (sheet, block, 3, i + block.startRowIndex, block.colCount) ;
					line.tpl = tpl ;
					line.colIndex = 3 ;
					line.iOption = GetLineInsertOption (RangeHelper.GetCell (sheet, 2, i + block.startRowIndex).Value2 as string) ;
					line.tplCellCount = block.colCount ;
					block.lineList.Add (line) ;
				}

				block.InitDColumn () ;
				if (block.dColumn != null)
					block.dColumn.tpl = tpl ;
				
				tpl.blockList.Add (block) ;
			}

			tpl.startRowIndex = lastValuedRowIndex + 5 ;
			return tpl ;
		}

		private static TplLine ParseLine (Worksheet sheet, TplBlock block, int startCol, int startRow, int colCount)
		{
			TplLine line = new TplLine () ;

			line.tplRange = RangeHelper.GetRange (sheet, startCol, startRow, colCount, 1) ;


			for (int colIndex = 0; colIndex < colCount; colIndex++)
			{
				Range range = RangeHelper.GetCell (sheet, startCol + colIndex, startRow) ;

				TplCell cell = new TplCell () ;

				cell.acrossColumns = 1 ;
				cell.tplRange = range ;
				cell.lastColIndex = colIndex + startCol ;

				
				string text = range.Value2 as string ;

				if (! string.IsNullOrEmpty(text))
					ParseCell (block, line, cell, text.Trim ()) ;
			
				line.cellList.Add(cell);
			}

			return line ;
		}

		private static void ParseCell (TplBlock block, TplLine line, TplCell cell, string text)
		{
			text = text.Trim () ;
			if (text.StartsWith ("R1C1:"))
			{
				cell.useR1C1Formula = true ;
				cell.tplTextContent = text.Substring (5) ;
				return ;
			}
			if (text [0] != '{')
			{
				// as text 
				cell.tplTextContent = text ;
				return ;
			}

			int i = text.IndexOf ('}') ;

			if (i > 0)
			{
				if (i + 1 != text.Length)
					cell.tplTextContent = text.Substring (i + 1) ;

				text = text.Substring (1, i - 1) ;
			}

			// using text as col name.
			cell.tplValueColName = text ;

			
			Dictionary<string, string> pair = ParseKeyValuePair (text) ;
			
			// parse format string
			cell.tplFormat = GetPairValue (pair, "f") ;
			if (! string.IsNullOrEmpty (cell.tplFormat))
				cell.tplFormat = cell.tplFormat.ToLower () ;
			
			// parse Merge option.
			cell.mOption = ParseMergeOption (GetPairValue (pair, "m")) ;
			
			// parse group options.
			if (GetPairValue (pair, "vg") != null)
			{
				cell.align = GroupAlign.vGroup ;

				cell.tplGroupColName = GetPairValue (pair, "vg").Trim ().ToUpper () ;
				
			}
			else if (GetPairValue (pair, "hg") != null)
			{
				cell.align = GroupAlign.hGroup ;
				cell.tplGroupColName = GetPairValue (pair, "hg").ToUpper () ;
				line.containsHGroup = true ;
				InsertOption option = GetLineInsertOption (GetPairValue (pair, "hgo")) ;
				if (option != InsertOption.afterChange && option != InsertOption.BeforeChange)
					option = InsertOption.afterChange;

				cell.hgOption = option ;
			}

			// parse value string including formula.
			string v = GetPairValue (pair, "v") ;

			if (string.IsNullOrEmpty (v))
			{
				if (!string.IsNullOrEmpty(cell.tplGroupColName))
					cell.tplValueColName = cell.tplGroupColName ;
				return ;
			} 
				

		    //pase default value
            cell.tplDefaultContent = GetPairValue(pair, "default");
		    
			i = v.IndexOf ('(') ;

			if (i < 0)
			{
				cell.tplValueColName = v.Trim ().ToUpper () ;
				return ;
			}

			cell.formula = new CellForumla () ;

			cell.formula.formulaName = i == 0 ? "" : v.Substring (0, i).ToLower () ;

			int i2 = v.IndexOf (')') ;
			
			if (i2 < 0)
			{
				Console.WriteLine ("Warning:: formula not closed. [" + v + "]");
				v = v.Substring (i + 1) ;
			}
			else 
				v = v.Substring (i + 1, i2 - i - 1) ;

			v = v.Trim () ;
			string[] colList = v.Split (new char[] {','}, StringSplitOptions.RemoveEmptyEntries) ;

			SearchKey pkey = null ;
			GroupValueSearchKey gkey = new GroupValueSearchKey () ;
			gkey.formula = cell.formula.formulaName ;
			cell.formula.keyList.Add (gkey) ;
			string colName = "" ;
			for (int j = 0; j < colList.Length; j++)
			{
				string colPair = colList [j] ;
				
				string[] cs = colPair.Split (new char[] {'='}, StringSplitOptions.RemoveEmptyEntries) ;

				if (j == 0)
				{
					colName = cs [0] ;
					/* gkey.formula = "" ; */
					gkey.valueColName = colName.Trim ().ToUpper () ;
					continue ;
				}

				

				SearchKey key = new SearchKey () ;
				key.colName = cs [0].ToUpper () ;
				
				if (cs.Length > 1)
				{
					key.isFixedValue = true ;
					key.keyValue = cs [1] ;
				}

				
				if (key.colName == "%")
				{
					key = GetGroupedKeys (block, line) ;
					
					if (key == null)
						continue ;
				}
		
				if (pkey == null)
				{
					gkey.key = key ;
					pkey = key ;
				}
				else
				{
					pkey.nextKey = key ;
					pkey = key ;
				
				}	
				while (pkey.nextKey != null)
				{
					pkey = pkey.nextKey ;
				}
			}

			block.gkeyList.Add (gkey.Copy ()) ;
		}

		private static SearchKey GetGroupedKeys (TplBlock block, TplLine line)
		{
			SearchKey root = null ;
			SearchKey pkey = null ;
			// get Line cell first 
			for (int i = 0; i < line.cellList.Count; i++)
			{
				TplCell leftCell = line.cellList [i] ;
				if (leftCell.align != GroupAlign.vGroup)
					continue ;
				SearchKey key = new SearchKey ();
				key.colName = leftCell.tplGroupColName ;
				if (root == null)
				{
					root = key ;
				}
				
				if (pkey == null)
				{
					pkey = key ;
				}
				else
				{
					pkey.nextKey = key ;
					pkey = key ;
				}
				
			}

			for (int i = 0; i < block.lineList.Count; i++)
			{
				TplLine l = block.lineList [i] ;
				
				if (l.cellList.Count <= line.cellList.Count)
					continue ;
				
				TplCell upCell = l.cellList [line.cellList.Count] ;
				if (upCell.align != GroupAlign.hGroup)
					continue ;

				SearchKey key = new SearchKey();
				key.colName = upCell.tplGroupColName;
				if (root == null)
				{
					root = key;
				}

				if (pkey == null)
				{
					pkey = key;
				}
				else
				{
					pkey.nextKey = key;
					pkey = key;
				}
				
			}
			return root ;
		}

		private static MergeOption ParseMergeOption (string text)
		{
			if (string.IsNullOrEmpty (text))
				return MergeOption.never ;

			
			switch (text.ToLower ())
			{
				case "up":
					return MergeOption.Up ;
				case "left":
					return MergeOption.Left ;
				default:
					return MergeOption.never ;
			}
		}

		private static string GetPairValue (Dictionary<string, string> pair, string key)
		{
			string value ;
			if (pair.TryGetValue (key, out value))
				return value ;
			else
				return null ;
		}


		private static InsertOption GetLineInsertOption (string option)
		{
			if (string.IsNullOrEmpty (option))
				return InsertOption.onfirst ;

			option = option.ToLower () ;

			switch (option)
			{
				case "last":
					return InsertOption.onLast ;
				case "before":
					return InsertOption.BeforeChange ;
				case "after":
					return InsertOption.afterChange ;
				case "never":
					return InsertOption.never ;
				case "all":
					return InsertOption.always ;
				default:
				case "first":
					return InsertOption.onfirst ;
			}
		}

		public static Dictionary<string, string> ParseKeyValuePair (string text)
		{
			Dictionary<string, string> map = new Dictionary<string, string> () ;
			string[] list = text.Split (new char[] {';'}, StringSplitOptions.RemoveEmptyEntries) ;


			for (int i = 0; i < list.Length; i++)
			{
				string str = list [i] ;

				string[] pair = str.Split (new char[] {':'}, StringSplitOptions.RemoveEmptyEntries) ;
				string value = "" ;
				if (pair.Length > 1)
					value = pair [1] ;
				// throw new ArgumentException ("String Format Error: [" + str + "]!") ;

				map.Add (pair [0].Trim ().ToLower (), value) ;
			}

			return map ;
		}
	}
}