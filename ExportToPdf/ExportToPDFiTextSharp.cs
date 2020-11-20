using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq; 
using iTextSharp.text;
using iTextSharp.text.pdf;


public class ExportToPDFiTextSharp
{
     
	string SelectedFont;
	float ReportTextSize = 7;

	float ColumnHeaderTextSize = 7;

	iTextSharp.text.BaseColor HeaderFontColor;
	private struct A4Dimension
	{
        //public const  Width1 = 842;
        //public const  Height1 = 595;
		//LandScape  (W x H, 842 x 595) points
		//Portrait (W x H, 595 × 842) points
	}

    public ExportToPDFiTextSharp( )
	{
		 
	}

	 
	/// <summary>
	/// 
	/// </summary>
	/// <param name="dt">Input Datatable to Export to PDF</param>
	/// <param name="strFilter">Filter conditions or remarks to display on report header</param>
	/// <param name="dicDataType">(ColumnIndex, XLDatattype)</param>
	/// <param name="lstRepeatColumn">Left most ColumnIndex to repeat on all pages</param>
	/// <param name="lstColumnsDisplay">ColumnIndex to display. Count = 0  or Nothing will Display All columns</param>
	/// <returns></returns>
	/// <remarks></remarks>
	public bool ExportToPDF(DataTable dt, string strFilter, Dictionary<int,Helper.Alignment> dicDataType, List<int> lstRepeatColumn = null, List<int> lstColumnsDisplay = null)
	{


		try {
			ICollection<string> myCol;// = ICollection<string>;
			//'/Returns the list of all font families included in iTextSharp.
			myCol = iTextSharp.text.FontFactory.RegisteredFamilies;
			//'Returns the list of all fonts included in iTextSharp.
			myCol = iTextSharp.text.FontFactory.RegisteredFonts;

			//FontFactory.Register("F:\Transfer\SEGOEUI.TTF")
			//SelectedFont = FontFactory.GetFont("SEGOEUI")

			// Dim Segoe As Font = FontFactory.GetFont("SegoeUI") ' , BaseFont.IDENTITY_H, 8)
			//Dim bfTimes As BaseFont = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, False)
			//Dim customfont As BaseFont
			//Try
			//    'customfont = BaseFont.CreateFont("F:\Transfer\SEGOEUI.TTF", BaseFont.CP1252, BaseFont.EMBEDDED)
			//    ' customfont = BaseFont.CreateFont(Segoe) '"F:\Transfer\SEGOEUI.TTF", BaseFont.CP1252, BaseFont.EMBEDDED)
			//    FontFactory.Register("F:\Transfer\SEGOEUI.TTF")
			//Catch ex As Exception
			//    Dim s = ex.Message
			//End Try


			int noOfColumns = dt.Columns.Count;
			int noOfRows = dt.Rows.Count;
			Dictionary<int, float> dicWidth = new Dictionary<int, float>();

			if ((lstRepeatColumn == null))
				lstRepeatColumn = new List<int>();
			if ((lstColumnsDisplay == null))
				lstColumnsDisplay = new List<int>();
			if ((dicDataType == null)) {
				dicDataType = new Dictionary<int, Helper.Alignment>();
				for (int icol = 0; icol <= noOfColumns - 1; icol++) {
					dicDataType.Add(icol, Helper.Alignment.Left);
				}
			}

            SelectedFont = "Arial";
			// "Arial"  SEGOEUI COURIER HELVETICA
			HeaderFontColor = iTextSharp.text.BaseColor.BLUE;

			List<iTextSharp.text.pdf.PdfPTable> lstTables = new List<iTextSharp.text.pdf.PdfPTable>();
			// Creates a PDF document
			Document document = null;
			//If LandScape = True Then
            //document = new Document( PageSize.A4.Rotate   , 0.0f, 0.0f, 40.0f, 15.0f);
           document=  new Document(new RectangleReadOnly(595,842,90) , 0.0f, 0.0f, 40.0f, 15.0f); 
			//LandScape  (W x H, 842 x 595) points
			//Else
			//'document = New Document(PageSize.A4, 0, 0, 15, 5) 'Portrait (W x H, 595 × 842) points
			//End If

			//dicWidth = GetColumnWidth(dt)
			dicWidth = GetColumnWidth(dt, SelectedFont, ColumnHeaderTextSize);

			Dictionary<int, List<int>> dicColsPerPage = new Dictionary<int, List<int>>();

			dicColsPerPage = GetColumnsPerPageByWidth(dt ,  dicWidth, lstRepeatColumn, lstColumnsDisplay);

			foreach (int intDt in dicColsPerPage.Keys) {
				List<int> lstColumnNums = dicColsPerPage[intDt];

				float[] arrColsrelativeWidths = new float[lstColumnNums.Count];
				dynamic dblTotalWidth = dicWidth.Sum(t => t.Value);

				int iArrIndex = 0;

				foreach (int iCol in lstColumnNums) {
					arrColsrelativeWidths[iArrIndex] = dicWidth[iCol];
					//dblTotalWidth = dblTotalWidth + dicWidth(iCol)
					iArrIndex = iArrIndex + 1;
				}

				iTextSharp.text.pdf.PdfPTable mainTable = new iTextSharp.text.pdf.PdfPTable(arrColsrelativeWidths);

				mainTable.TotalWidth = document.PageSize.Width * 0.9f;

				AddPageHeaderToPDFDataTable(ref mainTable, dt.TableName, strFilter, lstColumnNums.Count - 1);
				AddColumnHeader(dt, ref mainTable, lstColumnNums, dicDataType, lstColumnsDisplay);

				//'  If dblTotalWidth <= A4Dimension.Width * 0.9 - 30 Then mainTable.LockedWidth = True
				Phrase ph = default(Phrase);

				//' Date - centre, Double Right, String Left, INteger centre
				// Reads the gridview rows and adds them to the mainTable

				for (int rowNo = 0; rowNo <= noOfRows - 1; rowNo++) {
					foreach (int iCol in lstColumnNums) {
						if (lstColumnsDisplay.Count > 0 && lstColumnsDisplay.Contains(iCol) == false)
							continue;
						if (iCol == 999) {
							mainTable.AddCell(EmptyCell());
						} else {
							string sData = (dt.Rows[rowNo][iCol] == null) ? string.Empty : dt.Rows[rowNo][iCol].ToString().Trim();
							ph = new Phrase(sData, FontFactory.GetFont(SelectedFont, ReportTextSize, iTextSharp.text.Font.NORMAL));

							PdfPCell cell = new PdfPCell(ph);
							//cell.HorizontalAlignment =  GetAlignMent(dicDataType[iCol]);
                            if ((int)dicDataType[iCol] == (int)Helper.Alignment.Center) cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            if ((int)dicDataType[iCol] == (int)Helper.Alignment.Right) cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                            if ((int)dicDataType[iCol] == (int)Helper.Alignment.Left) cell.HorizontalAlignment = Element.ALIGN_LEFT;

                            if (lstRepeatColumn.Contains(iCol)) { cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW; }

							cell.NoWrap = true;
							cell.VerticalAlignment = Element.ALIGN_TOP;
							//Element.ALIGN_TOP
							mainTable.AddCell(cell);
						}
					}

					mainTable.CompleteRow();
					// Tells the mainTable to complete the row even if any cell is left incomplete.

					if (mainTable.TotalHeight >= 595-65 ){ // A4Dimension.Height - 65) {
						lstTables.Add(mainTable);

						mainTable = new iTextSharp.text.pdf.PdfPTable(arrColsrelativeWidths);
						mainTable.TotalWidth = document.PageSize.Width * 0.9f;

						AddPageHeaderToPDFDataTable(ref mainTable, dt.TableName, strFilter, lstColumnNums.Count - 1);
						AddColumnHeader(dt, ref mainTable, lstColumnNums, dicDataType, lstColumnsDisplay);

					}
				}

				mainTable.CompleteRow();
				lstTables.Add(mainTable);
			}
			//Dict           

			string strFileName = @"C:\Code\Project\TrailCode\eCatenate"  + "\\" + dt.TableName + "_" + DateTime.Now.ToString("yyyyMMdd HHmmss") + ".pdf";

			// Gets the instance of the document created and writes it to the output stream of the Response object.
			PdfWriter pdfWrite = PdfWriter.GetInstance(document, new FileStream(strFileName, FileMode.Create));
			// Response.OutputStream)
			pdfPage _pdfpage = new pdfPage(SelectedFont, HeaderFontColor);

			pdfWrite.PageEvent = _pdfpage;
			// Creates a footer for the PDF document.
			//Dim pdfFooter As New HeaderFooter(New Phrase("Page : ", FontFactory.GetFont(FontFactory.COURIER, FooterTextSize, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.DARK_GRAY)), True)

			//pdfFooter.Alignment = Element.ALIGN_CENTER
			//pdfFooter.Border = iTextSharp.text.Rectangle.NO_BORDER

			var _with1 = document;
			//.Footer = pdfFooter
			_with1.Open();

			//document.Add(New Paragraph(strFont, times))

			_with1.AddCreator("ElaiyaKumar");
			_with1.AddAuthor("By ElaiyaKumar");
			foreach (iTextSharp.text.pdf.PdfPTable oMainTable in lstTables) {
				_with1.Add(oMainTable);
				_with1.NewPage();
			}
			_with1.Close();

			document = null;
            Console.WriteLine("Any key to save ...."); 
            Console.ReadKey(); 
			//Interaction.MsgBox("Report Saved as " + strFileName, MsgBoxStyle.Information, "Export PDF");
            Console.WriteLine(strFileName);
            Console.WriteLine(); 

			Process.Start(strFileName);
			return true;
		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
		return false;
	}

	private PdfPCell EmptyCell()
	{
		PdfPCell cell = new PdfPCell(new Phrase());
		cell.Border = 0;
		return cell;
	}

    //private Element.al GetAlignMent(int intDataType) 
    //{
    //    try {
    //        //' Date - centre, Double Right, String Left, INteger centre
    //        if (intDataType == (int) Helper.Alignment.Center )
    //            return (int)Element.ALIGN_CENTER;

    //        if (intDataType == (int) Helper.Alignment.Right)
    //            return (int) Element.ALIGN_RIGHT;
    //        if (intDataType ==(int) Helper.Alignment.Left )
    //            return (int)Element.ALIGN_LEFT;

    //    } catch (Exception ex) {
    //        //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
    //    }
    //    return (int)Element.ALIGN_LEFT;
    //}

	private Dictionary<int, List<int>> GetColumnsPerPageByWidth(DataTable dt, Dictionary<int, float> dicWidth, List<int> lstRepeatColumn, List<int> lstColumnsDisplay)
	{
        Dictionary<int, List<int>> dicCol = new Dictionary<int, List<int>>();
		try {
			
            const float dblAllowedWidth = 842f * 0.9f - 30f;  //A4Dimension.Width * 0.9f - 30f;
			// 650
			double dblMaxWidth = 0;
			double dblWidth = 0;
			List<int> lstColumnNums = new List<int>();

			double dblRepeatColWidth = 0;
			foreach (int indx in lstRepeatColumn) {
				dblRepeatColWidth = dblRepeatColWidth + dicWidth[indx];
				//* 1.02
			}
			dblMaxWidth = dblAllowedWidth - dblRepeatColWidth;
			dblWidth = dblRepeatColWidth;

			for (int iCol = 0; iCol <= dt.Columns.Count - 1; iCol++) {
				if (lstColumnsDisplay.Count > 0 && lstColumnsDisplay.Contains(iCol) == false)
					continue;

				if (dblWidth + dicWidth[iCol] <= dblMaxWidth) {
					//lstColumnNums.Add(iCol)
					dblWidth = dblWidth + dicWidth[iCol];
				} else {
					dicCol.Add(dicCol.Count + 1, lstColumnNums);
					lstColumnNums = new List<int>();

					lstColumnNums.AddRange(lstRepeatColumn);
					dblWidth = dblRepeatColWidth;
				}
				lstColumnNums.Add(iCol);
			}
			dicCol.Add(dicCol.Count + 1, lstColumnNums);

			float  dblW =  (float) (dblMaxWidth - dblWidth);
			if (dblW > 10) {
				//lstColumnNums = dicCol(dicCol.Count)
				lstColumnNums.Add(999);
				//dicCol(dicCol.Count) = lstColumnNums
				dicWidth.Add(999, (dblW - 2f));
			}

			

		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
        return dicCol;
	}

	private bool AddPageHeaderToPDFDataTable(ref iTextSharp.text.pdf.PdfPTable mainTable, string sName, string strFilter, int noOfColumns)
	{

		try {
			//'mainTable.LockedWidth = True
			float ReportNameSize = 11;
			float ApplicationNameSize = 11;
			float FooterTextSize = 8;
			float HeaderTextSize = 9;

			float[] arColWidth1 = new float[3];
			arColWidth1[0] = 350;
			arColWidth1[1] = 300;
			arColWidth1[2] = 150;

			// Creates a PdfPTable with 3 columns to hold the header in the exported PDF.
			iTextSharp.text.pdf.PdfPTable HeaderTable = new iTextSharp.text.pdf.PdfPTable(arColWidth1);

			// Creates a phrase to hold the application name at the left hand side of the header.
			Phrase phApplicationName = new Phrase("Peacock", FontFactory.GetFont(SelectedFont, ApplicationNameSize, iTextSharp.text.Font.NORMAL));
			phApplicationName.Font.Color = HeaderFontColor;
			// iTextSharp.text.BaseColor.BLUE

			PdfPCell clApplicationName = new PdfPCell(phApplicationName);
			clApplicationName.Border = PdfPCell.NO_BORDER;
			clApplicationName.HorizontalAlignment = Element.ALIGN_LEFT;

			// Creates a phrase to show the current date at the right hand side of the header.
            Phrase phDate = new Phrase(DateTime.Now.ToString("dd-MM-yy HH:mm"), FontFactory.GetFont(SelectedFont, HeaderTextSize, iTextSharp.text.Font.NORMAL));
			phDate.Font.Color = HeaderFontColor;
			//iTextSharp.text.BaseColor.BLUE

			PdfPCell clDate = new PdfPCell(phDate);
			clDate.HorizontalAlignment = Element.ALIGN_RIGHT;
			clDate.Border = PdfPCell.NO_BORDER;

			Phrase phName = new Phrase(sName, FontFactory.GetFont(SelectedFont, HeaderTextSize, iTextSharp.text.Font.NORMAL));
			phName.Font.Color = HeaderFontColor;
			//iTextSharp.text.BaseColor.BLUE

			PdfPCell cellName = new PdfPCell(phName);
			cellName.HorizontalAlignment = Element.ALIGN_LEFT;
			cellName.Border = PdfPCell.NO_BORDER;

			HeaderTable.AddCell(cellName);
			HeaderTable.AddCell(clApplicationName);
			HeaderTable.AddCell(clDate);


			Phrase phFilter = new Phrase(strFilter, FontFactory.GetFont(SelectedFont, HeaderTextSize, iTextSharp.text.Font.NORMAL));
			phFilter.Font.Color = HeaderFontColor;
			//iTextSharp.text.BaseColor.BLUE
			PdfPCell clFilter = new PdfPCell(phFilter);
			clFilter.Colspan = 2;
			clFilter.Border = PdfPCell.NO_BORDER;
			clFilter.HorizontalAlignment = Element.ALIGN_LEFT;


			Phrase phUserid = new Phrase("Elaiya Kumar", FontFactory.GetFont(SelectedFont, HeaderTextSize, iTextSharp.text.Font.NORMAL));
			phUserid.Font.Color = HeaderFontColor;
			//iTextSharp.text.BaseColor.BLUE
			PdfPCell clUserid = new PdfPCell(phUserid);
			clUserid.HorizontalAlignment = Element.ALIGN_RIGHT;
			clUserid.Border = PdfPCell.NO_BORDER;

			HeaderTable.AddCell(clFilter);
			//rowTabe1.AddCell(EmptyCell())
			HeaderTable.AddCell(clUserid);

			// Creates a PdfPCell that accepts the headerTable as a parameter and then adds that cell to the main PdfPTable.
			PdfPCell cellHeader1 = new PdfPCell(HeaderTable);
			cellHeader1.Border = PdfPCell.NO_BORDER;
			cellHeader1.Colspan = noOfColumns + 1;
			mainTable.AddCell(cellHeader1);


			//mainTable.AddCell(clph)
			Phrase ph = new Phrase();
			// Creates a phrase for a new line.
            Phrase phSpace = new Phrase("\r\n");
			phSpace.Font.Size = 1;

			PdfPCell clSpace = new PdfPCell(phSpace);
			clSpace.Border = PdfPCell.NO_BORDER;
			clSpace.Colspan = noOfColumns + 1;
			// noOfColumns
			mainTable.AddCell(clSpace);

			mainTable.CompleteRow();

			return true;

		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
        return true;
	}


	private bool AddPageHeaderToPDFDataTableLogo(ref iTextSharp.text.pdf.PdfPTable mainTable, string sName, int noOfColumns)
	{
		//As iTextSharp.text.pdf.PdfPTable

		try {
            string strPath = ""; //My.Application.Info.DirectoryPath + "\\syneco_rgb.jpg"
			iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(strPath);
			//'mainTable.LockedWidth = True
			float ReportNameSize = 11;
			float ApplicationNameSize = 11;
			float FooterTextSize = 8;
			float HeaderTextSize = 9;


			// Creates a PdfPTable with 3 columns to hold the header in the exported PDF.
			iTextSharp.text.pdf.PdfPTable headerTable = new iTextSharp.text.pdf.PdfPTable(3);

			// Creates a phrase to hold the application name at the left hand side of the header.
			Phrase phApplicationName = new Phrase("Eagle", FontFactory.GetFont(SelectedFont, ApplicationNameSize, iTextSharp.text.Font.NORMAL));
			//phApplicationName.Font = SelectedFont
			//'phApplicationName.Font.SetFamily(iTextSharp.text.Font.FontFamily.COURIER) ' .Font.SetFamily(SelectedFont)
			//'With phApplicationName.Font
			//'    .Size = ApplicationNameSize
			//'    '.SetFamily(iTextSharp.text.Font.FontFamily.COURIER)
			//'    .SetStyle(iTextSharp.text.Font.NORMAL)
			//'End With
			// Creates a PdfPCell which accepts a phrase as a parameter.
			PdfPCell clApplicationName = new PdfPCell(phApplicationName);

			// Sets the border of the cell to zero.
			clApplicationName.Border = PdfPCell.NO_BORDER;

			// Sets the Horizontal Alignment of the PdfPCell to left.
			clApplicationName.HorizontalAlignment = Element.ALIGN_CENTER;

			// Creates a phrase to show the current date at the right hand side of the header.
			Phrase phDate = new Phrase(DateTime.Now.Date.ToString("dd-MMM-yyyy"), FontFactory.GetFont(SelectedFont, HeaderTextSize, iTextSharp.text.Font.NORMAL));

			// Creates a PdfPCell which accepts the date phrase as a parameter.
			PdfPCell clDate = new PdfPCell(phDate);

			// Sets the Horizontal Alignment of the PdfPCell to right.
			clDate.HorizontalAlignment = Element.ALIGN_RIGHT;

			// Sets the border of the cell to zero.
			clDate.Border = PdfPCell.NO_BORDER;
			PdfPCell cllogo = new PdfPCell(logo);

			cllogo.HorizontalAlignment = Element.ALIGN_LEFT;

			cllogo.Border = PdfPCell.NO_BORDER;
			headerTable.AddCell(cllogo);

			// Adds the cell which holds the application name to the headerTable.
			headerTable.AddCell(clApplicationName);

			// Adds the cell which holds the date to the headerTable.
			headerTable.AddCell(clDate);

			// Creates a phrase which holds the file name.
			Phrase phHeader = new Phrase(sName, FontFactory.GetFont(SelectedFont, ReportNameSize, iTextSharp.text.Font.BOLD));
			//phHeader.Font = SelectedFont
			//'phHeader.Font.SetFamily(iTextSharp.text.Font.FontFamily.COURIER) '.Font.SetFamily(SelectedFont)
			//'With phHeader.Font
			//'    .Size = ReportNameSize
			//'    '.SetFamily(iTextSharp.text.Font.FontFamily.COURIER)
			//'    .SetStyle(iTextSharp.text.Font.BOLD)
			//'End With

			PdfPCell clHeader = new PdfPCell(phHeader);
			clHeader.Colspan = noOfColumns;
			clHeader.Border = PdfPCell.NO_BORDER;
			clHeader.HorizontalAlignment = Element.ALIGN_CENTER;
			headerTable.AddCell(clHeader);

			// Dim phFilter As New Phrase(strFormulaToPrint)

			//With phFilter.Font
			//    .Size = HeaderTextSize ' ReportTextSize
			//    .SetFamily(iTextSharp.text.Font.FontFamily.COURIER)
			//    .SetStyle(iTextSharp.text.Font.NORMAL)
			//End With

			//Dim clFilter As New PdfPCell(phFilter)
			//clFilter.Colspan = noOfColumns
			//clFilter.HorizontalAlignment = Element.ALIGN_LEFT
			//clFilter.Border = PdfPCell.NO_BORDER
			//headerTable.AddCell(clFilter)

			// Creates a PdfPCell that accepts the headerTable as a parameter and then adds that cell to the main PdfPTable.
			PdfPCell cellHeader = new PdfPCell(headerTable);
			cellHeader.Border = PdfPCell.NO_BORDER;

			// Sets the column span of the header cell to noOfColumns.
			cellHeader.Colspan = noOfColumns;

			mainTable.AddCell(cellHeader);

			//mainTable.AddCell(clph)
			Phrase ph = new Phrase();
			// Creates a phrase for a new line.
            Phrase phSpace = new Phrase("\r\n");
			phSpace.Font.Size = 1;

			PdfPCell clSpace = new PdfPCell(phSpace);
			clSpace.Border = PdfPCell.NO_BORDER;
			clSpace.Colspan = noOfColumns;
			// noOfColumns
			mainTable.AddCell(clSpace);

			mainTable.CompleteRow();
			return true;

		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
        return true;
	}


	private bool AddColumnHeader(DataTable dt, ref iTextSharp.text.pdf.PdfPTable mainTable, List<int> lstColumnNums, Dictionary<int, Helper.Alignment> dicDataType, List<int> lstColumnsDisplay)
	{
		try {
			mainTable.DefaultCell.BackgroundColor = iTextSharp.text.BaseColor.BLACK;

			// Sets the gridview column names as table headers.

			foreach (int iCol in lstColumnNums) {
				if (lstColumnsDisplay.Count > 0 && lstColumnsDisplay.Contains(iCol) == false)
					continue;

				Phrase ph = default(Phrase);
				//Dim strText As String = If(iCol = 999, String.Empty, dt.Columns(iCol).Caption)
				if (iCol == 999) {
					PdfPCell cell = EmptyCell();
					cell.BackgroundColor = iTextSharp.text.BaseColor.WHITE;
					mainTable.AddCell(cell);
				} else {
					ph = new Phrase(dt.Columns[iCol].Caption, FontFactory.GetFont(SelectedFont, ColumnHeaderTextSize, iTextSharp.text.Font.NORMAL));
					ph.Font.Color = iTextSharp.text.BaseColor.WHITE;

					PdfPCell cell = new PdfPCell(ph);
					//cell.HorizontalAlignment = GetAlignMent(dicDataType[iCol]);

                    if ((int)dicDataType[iCol] == (int)Helper.Alignment.Center) cell.HorizontalAlignment = Element.ALIGN_CENTER;
                    if ((int)dicDataType[iCol] == (int)Helper.Alignment.Right) cell.HorizontalAlignment =  Element.ALIGN_RIGHT;
                    if ((int)dicDataType[iCol] == (int)Helper.Alignment.Left) cell.HorizontalAlignment = Element.ALIGN_LEFT;

					cell.BackgroundColor = iTextSharp.text.BaseColor.GREEN;
					//cell. Color == iTextSharp.text.BaseColor.BLACK
					cell.NoWrap = true;
					mainTable.AddCell(cell);
				}
			}
			mainTable.CompleteRow();
			mainTable.DefaultCell.BackgroundColor = iTextSharp.text.BaseColor.WHITE;

		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
        return true;
	}

	private Dictionary<int, float> GetColumnWidth(DataTable dt, string FontName, float FontSize)
	{
        Dictionary<int, float> dicW = new Dictionary<int, float>();

		try {
			System.Drawing.Font fnt = new System.Drawing.Font(FontName, FontSize, FontStyle.Regular, GraphicsUnit.Point);

			
			//Using graphics As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(New Bitmap(1, 1))
			//    Dim size As SizeF = graphics.MeasureString(Str, New System.Drawing.Font(sFontName, 7, FontStyle.Regular, GraphicsUnit.Point))
			//    Debug.Print(sFontName & "; " & Str() & "; " & size.Width.ToString)
			//End Using

			// Compute the string dimensions in the given font         
			using ( System.Drawing.Graphics oGrapImage = Graphics.FromImage(new Bitmap(1, 1))) {

				for (int iCol = 0; iCol <= dt.Columns.Count - 1; iCol++) {
					SizeF stringSize = oGrapImage.MeasureString(dt.Columns[iCol].Caption.ToString(), fnt);
					dicW.Add(dicW.Count, stringSize.Width);
				}

				for (int iRow = 0; iRow <= dt.Rows.Count - 1; iRow++) {
					for (int iCol = 0; iCol <= dt.Columns.Count - 1; iCol++) {
						SizeF stringSize = oGrapImage.MeasureString(dt.Rows[iRow][iCol].ToString(), fnt);

						if (stringSize.Width > dicW[iCol])
							dicW[iCol] = stringSize.Width * 1.5f;
					}
				}

			}

		

		} catch (Exception ex) {
            //PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
		}
        	return dicW;
	}

    //private bool GetColumnWidth()
    //{

    //    try {
    //        List<string> lst = new List<string> {
    //            "",
    //            "W",
    //            "O",
    //            "p",
    //            "02/07/2015",
    //            "Hour 12-13",
    //            "Hour 20-24",
    //            "12.39",
    //            "12.3912.39",
    //            "What",
    //            "WWWW"
    //        };
    //        string sFontName = "Arial";
    //        "COURIER"

    //        using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(new Bitmap(1, 1))) {
    //            foreach (string Str in lst) {
    //                SizeF size = graphics.MeasureString(Str, new System.Drawing.Font(sFontName, 7, FontStyle.Regular, GraphicsUnit.Point));
    //                Debug.Print(sFontName + "; " + Str + "; " + size.Width.ToString);
    //            }

    //        }

    //    } catch (Exception ex) {
    //        PHLog.ErrorLogException(ex, oGV, System.Reflection.MethodBase.GetCurrentMethod.Name);
    //    }

    //}

	private class pdfPage : iTextSharp.text.pdf.PdfPageEventHelper
	{
		string SelectedFont;
		iTextSharp.text.BaseColor FontColor;
		public pdfPage(string _SelectedFont, iTextSharp.text.BaseColor _FontColor)
		{
			SelectedFont = _SelectedFont;
			FontColor = _FontColor;
		}

		public override void OnEndPage(PdfWriter writer, Document doc)
		{
			//I use a PdfPtable with 2 columns to position my footer where I want it
			PdfPTable footerTbl = new PdfPTable(1);

			//set the width of the table to be the same as the document
			footerTbl.TotalWidth = doc.PageSize.Width;

			//Center the table on the page
			footerTbl.HorizontalAlignment = Element.ALIGN_CENTER;

			//Create a paragraph that contains the footer text
			Phrase ph = new Phrase("Page " + writer.PageNumber.ToString(), FontFactory.GetFont(SelectedFont, 8, iTextSharp.text.Font.NORMAL));
			ph.Font.Color = FontColor;
			//ph.Font.SetFamily(SelectedFont)
			//'With ph.Font
			//'    .Size = 8
			//'    '.SetFamily(iTextSharp.text.Font.FontFamily.COURIER)
			//'    .SetStyle(iTextSharp.text.Font.NORMAL)
			//'End With

			//create a cell instance to hold the text
			PdfPCell cell = new PdfPCell(ph);
			cell.Border = 0;

			cell.HorizontalAlignment = Element.ALIGN_CENTER;
			footerTbl.AddCell(cell);
			footerTbl.WriteSelectedRows(0, 1, 0, (doc.BottomMargin - 2), writer.DirectContent);
		}
	}
}
 