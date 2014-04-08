using System;
using Microsoft.Office.Interop.Excel;

namespace ReservoirSimulator.Utilities
{
	public class ExcelDoc
	{
        
		private readonly Application _app;
		private readonly Workbook _workbook;
		private readonly Worksheet _worksheet;
		
		private Range _workSheetRange;

		public ExcelDoc()
		{
			_app = new Application { Visible = true };
			_workbook = _app.Workbooks.Add(1);
			_worksheet = (Worksheet) _workbook.Sheets[1];
		}

		/// <summary>
		/// Creates a header in the excel sheet
		/// </summary>
		/// <param name="row">the row of the header</param>
		/// <param name="col">the column of the header</param>
		/// <param name="htext">the text to be shown in the header</param>
		/// <param name="cell1">starting cell range</param>
		/// <param name="cell2">ending cell range</param>
		/// <param name="mergeColumns"></param>
		/// <param name="interiorColor"></param>
		/// <param name="isFontBold"></param>
		/// <param name="fontSize"></param>
		/// <param name="fontColor"></param>
		public void CreateHeaders(int row, int col, string htext, string cell1, string cell2, int mergeColumns, string interiorColor, bool isFontBold, int fontSize, string fontColor)
		{
			_worksheet.Cells[row, col] = htext;
			_workSheetRange = _worksheet.Range[cell1, cell2];
			_workSheetRange.Merge(mergeColumns);

			switch (interiorColor)
			{
				case "YELLOW":
					_workSheetRange.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
					break;
				case "GRAY":
					_workSheetRange.Interior.Color = System.Drawing.Color.Gray.ToArgb();
					break;
				case "GAINSBORO":
					_workSheetRange.Interior.Color = 
						System.Drawing.Color.Gainsboro.ToArgb();
					break;
				case "Turquoise":
					_workSheetRange.Interior.Color = 
						System.Drawing.Color.Turquoise.ToArgb();
					break;
				case "PeachPuff":
					_workSheetRange.Interior.Color = 
						System.Drawing.Color.PeachPuff.ToArgb();
					break;
				default:
					//  workSheet_range.Interior.Color = System.Drawing.Color..ToArgb();
					break;
			}
         
			_workSheetRange.Borders.Color = System.Drawing.Color.Black.ToArgb();
			_workSheetRange.Font.Bold = isFontBold;
			_workSheetRange.ColumnWidth = fontSize;
			_workSheetRange.Font.Color = fontColor.Equals("") 
				? System.Drawing.Color.White.ToArgb() 
				: System.Drawing.Color.Black.ToArgb();
		}

		/// <summary>
		/// Add data to a particular cell
		/// </summary>
		/// <param name="row"></param>
		/// <param name="col"></param>
		/// <param name="data"></param>
		public void AddData(int row, int col, double data)
		{
			_worksheet.Cells[row, col] = data;
		}   

	}
}