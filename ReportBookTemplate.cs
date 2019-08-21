using System;
using System.Collections.Generic;
using System.Data ;
using System.Text;
using AFC.WorkStation.DB ;
using Spire.Xls ;

namespace AFC.WorkStation.ExcelReport
{
	public class ReportBookTemplate
	{
		public List<ReportSheetTemplate> sheetList = new List<ReportSheetTemplate> ();
		public DataSet dataSet = new DataSet ();


		public void LoadTemplate (DBO db, Workbook book, Dictionary<string,object> paramMap)
		{
			for (int i = 0; i < book.Worksheets.Count; i++)
			{
				Worksheet sheet = book.Worksheets [i] ;

				ReportSheetTemplate tpl = TplLoader.ParseTemplate(sheet, i + 1);

				sheetList.Add(tpl);
				/* tpl.startRowIndex = 300; */
				tpl.dset = dataSet;
				tpl.paramMap = paramMap;

				tpl.LoadDataSource(db);
			}

			for (int i = 0; i < sheetList.Count; i++)
			{
				ReportSheetTemplate tpl = sheetList [i] ;

				for (int j = 0; j < tpl.blockList.Count; j++)
				{
					TplBlock block = tpl.blockList [j] ;
					if (! block.isChart)
						continue ;
					
					if (block.chartDataBlock != null)
						continue ;
						
					// find block in each sheet.
					block.chartDataBlock = FindDataBlock(block.chartDataBlockName);
				}
			}
			return ;
		}

		private TplBlock FindDataBlock (string blockName)
		{
			for (int i = 0; i < sheetList.Count; i++)
			{
				ReportSheetTemplate tpl = sheetList [i] ;

				for (int j = 0; j < tpl.blockList.Count; j++)
				{
					TplBlock block = tpl.blockList [j] ;
					if (block.name == blockName)
						return block ;
				}
			}
			
			return null ;
		}

		public void FillTemplate ()
		{
			DateTime now = DateTime.Now ;
			for (int i = 0; i < sheetList.Count; i++)
			{
				ReportSheetTemplate tpl = sheetList [i] ;
				tpl.FillTemplate();
			}

			// setup chart
			for (int i = 0; i < sheetList.Count; i++)
			{
				ReportSheetTemplate tpl = sheetList [i] ;

				for (int j = 0; j < tpl.blockList.Count; j++)
				{
					TplBlock block = tpl.blockList [j] ;
					if (! block.isChart)
						continue ;
					
					block.SetupChart () ;
				}
			}

			for (int i = 0; i < sheetList.Count; i++)
			{
				ReportSheetTemplate tpl = sheetList[i];

				for (int j = 0; j < tpl.blockList.Count; j++)
				{
					TplBlock block = tpl.blockList[j];
					
					Console.WriteLine ("Block: " + block.name + 
						" gKeyList: " + block.gkeyList.Count +
						" ValueList: " + block.holder.valueList.Count);
				}
			}

			Console.WriteLine("GetColVlaue Call Times: " + RangeHelper.getColValueCount);
			Console.WriteLine("GetColVlaueByIndex Call Times: " + RangeHelper.getColValueByIndexCount);
			Console.WriteLine("getReusedValueCallCount Call Times: " + ReusedKey.getReusedValueCallCount);
			Console.WriteLine("columnNameFindCount Call Times: " + ReusedKey.columnNameFindCount);
			Console.WriteLine("valueReusedCount Call Times: " + ReusedKey.valueReusedCount);
			Console.WriteLine("ReusedKeyPool Size: " + ReusedKey.keyPool.Count);
			
			Console.WriteLine ("GetRange Call Times: " + RangeHelper.getRangeCallTimes);
			Console.WriteLine ("UpdateCellValue Call Times: " + RangeHelper.getRangeCallTimes);
			Console.WriteLine("insertCopyRange Call Times: " + RangeHelper.insertCopyRangeCallTimes);
			Console.WriteLine ("Total time: " + (DateTime.Now.Subtract (now)).TotalMilliseconds + "ms.");
		}
	}
}
