using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace Formularz_FDC
{
	class FdcFormTemplate
	{
		String sheetName = "Formularz do druku";
		String sheetPassword = "NcMOi7X1fpLh";
		List<int> defaultMeasures = new List<int>()
										{
											19,
											24
										};
		List<int> columnsWidth = new List<int>()
										{
											0,
											0,
											6,
											24,
											6,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											24,
											6,
											0,
											0,
											//200
										};
		List<int> sectionCHeights = new List<int>()
										{
											9,
											6,
											14,
											14,
											14,
											14,
											14,
											14,
											14,
											14,
											14,
											14,
											19,
											19,
											19,
											14,
											19,
											14,
											14,
											3,
											14,
											14,
											19,
											19,
											19,
											14,
											19,
											14,
										};
		List<string> contentSectionC = new List<string>()
										{
											"OŚWIADCZENIE OSOBY ZGŁASZAJĄCEJ DANE",
											"1. treść oświadczenia",
											"1",
											"Oświadczam, że dane zawarte w formularzu są zgodne ze stanem prawnym i faktycznym.",
											"2",
											"Jestem świadomy/a odpowiedzialności za poświadczanie nieprawdy lub zatajenie prawdy.",
											"3",
											"W związku z wymogiem aktualizacji danych zgłaszanych do ubezpieczenia zdrowotnego zobowiązuję się podać do wiadomości ",
											"płatnika składek (pracodawcy) każdą zmianę danych o członkach rodziny, objętych niniejszym zgłoszeniem, ",
											"w ciągu 7 dni od jej wystąpienia.",
											"4",
											"Jestem poinformowany, że opóźnienie zgłoszenia zmiany danych ubezpieczonego naraża płatnika składek na uchybienie ",
											"ustawowym terminom zgłoszeń do ZUS.",
											"4. podpis osoby zgłaszającej dane",
											"2. miejsce oświadczenia (miejscowość)",
											"3. data oświadczenia (dd/mm/rrrr)",
											"ADNOTACJE KOMÓRKI KADROWEJ",
											"3. oznaczenie i podpis osoby przyjmującej",
											"4. oznaczenie podmiotu / pieczęć",
											"1. miejscowość",
											"2. data wpływu formularza (dd/mm/rrrr)"
										};
		List<int> sectionBHeights = new List<int>()
										{
											9,
											19,
											12,
											19,
											12,
											19,
											12,
											19,
											12,
											19,
											6,
											12,
											19,
											19,
											19,
											19,
											19,
											19,
											19,
											19,
											19,
											19,
											19,
											6,
											12,
											19,
											19,
											19,
											19,
											19,
											14,
											14,
											12,
											19,
											12,
											19,
											12,
											19,
											12,
											19,
											12,
											19
										};
		List<string> contentSectionB = new List<string>()
										{
											"DANE O CZŁONKU RODZINY PRACOWNIKA - OSOBY UBEZPIECZONEJ, UPRAWNIONYM DO ŚWIADCZEŃ Z UBEZPIECZENIA ZDROWOTNEGO",
											". A. DANE O CZŁONKU RODZINY, UPRAWNIONYM DO ŚWIADCZEŃ Z UBEZPIECZENIA ZDROWOTNEGO",
											"1. imię pierwsze",
											"2. nazwisko ",
											"3. PESEL",
											"4. data urodzenia (dd/mm/rrrr)",
											"5. pozostaje we wspólnym gospodarstwie (tak/nie)",
											"6. seria i nr dokumentu tożsamości",
											"7. rodzaj dokumentu (dow.os./paszport)",
											"8. kod",
											"9. kod",
											"10. pokrewieństwo",
											"małżonek",
											"dziecko własne, przysposobione lub dziecko małżonka",
											"wnuk albo dziecko obce, dla którego ustanowiono opiekę, albo dziecko obce w ramach rodziny zastępczej lub rodzinnego domu dziecka",
											"matka",
											"ojciec",
											"macocha",
											"ojczym",
											"babka",
											"dziadek",
											"osoby przysposabiające osoby ubezpieczone",
											"inni wstępni pozostający z ubezpieczonym we współnym gospodarstwie domowym",
											"11. kod",
											"12. niepełnosprawność",
											"nie posiada orzeczenia o niepełnosprawności",
											"posiada orzeczenie o lekkim stopniu niepełnosprawności",
											"posiada orzeczenie o umiarkowanym stopniu niepełnosprawności",
											"posiada orzeczenie o znacznym stopniu niepełnosprawności",
											"posiada orzeczenie o niepełnosprawności wydawane osobom do 16 roku życia",
											". B. ADRES ZAMIESZKANIA ",
											"(wpisać, jeśli adres zamieszkania zgłaszanego członka rodziny jest inny, niż adres zamieszkania ubezpieczonego)",
											"1. województwo",
											"2. gmina",
											"3. miejscowość",
											"4. kod pocztowy",
											"5. ulica",
											"6. nr domu",
											"7. nr lokalu",
											"8. numer telefonu"
										};
		List<string> contentSectionBAlt = new List<string>()
										{
											"",
											". A. DANE O CZŁONKU RODZINY, UPRAWNIONYM DO ŚWIADCZEŃ Z UBEZPIECZENIA ZDROWOTNEGO",
											"",
										};
		List<int> sectionBHeightsAlt = new List<int>()
										{
											9,
											19,
											12,
											19
										};
		List<int> sectionAHeights = new List<int>()
										{
											12,
											19,
											14,
											19,
											4,
											19,
											12,
											19,
											12,
											19,
											12,
											19
										};
		List<string> contentSectionA = new List<string>()
										{
											"WYPEŁNIAĆ W WYZNACZONYCH KRATKACH KOMPUTEROWO LUB RĘCZNIE DRUKOWANYMI LITERAMI",
											"FDC",
											"FORMULARZ DANYCH CZŁONKA RODZINY ZGŁASZANEGO DO UBEZPIECZENIA ZDROWOTNEGO",
											"data wypełnienia formularza (dd/mm/rrrr)",
											"nr kontrolny",
											"DANE PRACOWNIKA",
											"I. DANE IDENTYFIKACYJNE OSOBY UBEZPIECZONEJ (PRACOWNIKA)",
											"1. nazwisko",
											"2. cd. nazwisko",
											"3. imię pierwsze"
										};
		List<string> relationshipCodes = new List<string>() { "01", "11", "21", "30", "31", "32", "33", "40", "41", "50", "60" };
		List<string> disabilityCodes = new List<string>() {"0", "1", "2", "3", "4" };
		public IWorkbook book;
		public int relativesCount;

		public FdcFormTemplate(string filepath, int relativesCount)
		{
			try {
				book = createTemplate(filepath, relativesCount);
				book.GetSheetAt(0).ProtectSheet(sheetPassword);
				using (var fs = new FileStream(filepath, FileMode.Create)) {
					book.Write(fs);
				}
			}
			catch (Exception ex) {
				Console.WriteLine(ex);
			}
		}

		public FdcFormTemplate(string filepath, FdcSectionA fdcSectionA, List<FdcSectionB> fdcSectionBList, FdcSectionC fdcSectionC)
		{
			try {
				book = createTemplate(filepath, fdcSectionBList.Count);
				book = fillTemplate(book, fdcSectionA, fdcSectionBList, fdcSectionC);
				book.GetSheetAt(0).ProtectSheet(sheetPassword);
				using (var fs = new FileStream(filepath, FileMode.Create)) {
					book.Write(fs);
				}

			}
			catch (Exception ex) {
				Console.WriteLine(ex);
			}

		}

		public IWorkbook fillTemplate( IWorkbook book, FdcSectionA fdcSectionA, List<FdcSectionB> fdcSectionB, FdcSectionC fdcSectionC)
		{
			var rn = 0;

			var sheet = book.GetSheetAt(0);

			//A
			rn += 3;
			WriteContent( sheet, fdcSectionA.date.ToCharArray(), rn, 5);
			WriteContent( sheet, fdcSectionA.id.ToCharArray(), rn, 29);

			rn += 4;
			WriteContent( sheet, fdcSectionA.surname.ToCharArray(), rn, 5, 29);

			rn += 4;
			WriteContent( sheet, fdcSectionA.name.ToCharArray(), rn, 5);

			if (fdcSectionB.Count != 0)
			{
				//B
				for (int i = 0; i < fdcSectionB.Count; i++)
				{
					rn += 4;
					WriteContent(sheet, fdcSectionB[i].name.ToCharArray(), rn, 5);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].surname.ToCharArray(), rn, 5);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].pesel.ToCharArray(), rn, 5);
					WriteContent(sheet, fdcSectionB[i].birth.ToCharArray(), rn, 17);
					WriteContent(sheet, fdcSectionB[i].commonGround.ToCharArray(), rn, 26);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].idId.ToCharArray(), rn, 5);
					WriteContent(sheet, fdcSectionB[i].idType.ToCharArray(), rn, 15);
					WriteContent(sheet, fdcSectionB[i].idNum.ToCharArray(), rn, 24);

					rn += 3;
					rn = WriteChecker(sheet, fdcSectionB[i].relationshipId, relationshipCodes, 5, 6, rn, relationshipCodes.Count);

					rn += 2;
					rn = WriteChecker(sheet, fdcSectionB[i].disabilityId, disabilityCodes, 5, 6, rn, disabilityCodes.Count);

					rn += 3;
					WriteContent(sheet, fdcSectionB[i].voivodeship.ToCharArray(), rn, 5);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].community.ToCharArray(), rn, 5);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].place.ToCharArray(), rn, 5);
					WriteContent(sheet, fdcSectionB[i].postal.ToCharArray(), rn, 28);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].street.ToCharArray(), rn, 5);

					rn += 2;
					WriteContent(sheet, fdcSectionB[i].flatNum.ToCharArray(), rn, 5);
					WriteContent(sheet, fdcSectionB[i].apartmentNum.ToCharArray(), rn, 14);
					WriteContent(sheet, fdcSectionB[i].telephone.ToCharArray(), rn, 20);
				}
			}
			else { rn += sectionBHeightsAlt.Count; }

			rn += 15;
			WriteContent( sheet, fdcSectionC.place.ToCharArray(), rn, 5, 22);

			rn += 2;
			WriteContent( sheet, fdcSectionC.date.ToCharArray(), rn, 5, 12);

			return book;
		}

		static int WriteChecker( ISheet sheet, string relationshipId, List<String>breakerR, int codeStart, int checkerStart, int row, int checkersCount)
		{
			int checker = -1;
			for (int i = 0; i < breakerR.Count; i++)
			{
				if (relationshipId == breakerR[i])
				{
					checker = i;
					break;
				} else {
					checker = -1;
				}
			}

			if (checker != -1) {
				row += checker;

				WriteCell(sheet, codeStart, row, relationshipId);
				WriteContent(sheet, "X".ToCharArray(), row, checkerStart);
				int toLast = 0;
				for (int i = 0; i < breakerR.Count; i++)
				{
					if (checker == i) {
						toLast = breakerR.Count - i;
						break;
					}
				}

				return row + toLast;
			} else {
				return checkersCount;
			}
		}

		public static void WriteContent( ISheet sheet, char[] sectionAPart, int row, int startLine)
		{
			for (int i = 0; i < sectionAPart.Length; i++)
			{
				WriteCell(sheet, startLine++, row, sectionAPart[i].ToString());
			}
		}

		public static void WriteContent( ISheet sheet, char[] sectionAPart, int row, int startLine, int breakLine)
		{
			for (int i = 0; i < sectionAPart.Length; i++)
			{
				if(i == breakLine)
				{
					row++;
					row++;
					startLine = 5;
				}
				WriteCell(sheet, startLine++, row, sectionAPart[i].ToString());
			}
		}

		public IWorkbook createTemplate(string filepath, int relativesCount)
		{
			//Zdefiniowanie skoroszytu
			var book = CreateNewBook(filepath);
			var sheet = book.CreateSheet(sheetName);
			sheet.DefaultRowHeight = (short)defaultMeasures[0];
			sheet.DefaultColumnWidth = (short)defaultMeasures[1];
			sheet.PrintSetup.PaperSize = (short)PaperSize.A4_Small;
			sheet.PrintSetup.Scale = 65;
			sheet.FitToPage = true;
			sheet.PrintSetup.FitWidth = (short)1;
			sheet.PrintSetup.FitHeight = (short)0;
			sheet.SetMargin(MarginType.RightMargin, 0.4015748031);
			sheet.SetMargin(MarginType.LeftMargin, 0.4015748031);
			sheet.SetMargin(MarginType.TopMargin, 0.3937007874);
			sheet.SetMargin(MarginType.BottomMargin, 0.3937007874);

			IFooter footer = sheet.Footer;
			footer.Center = HeaderFooter.Page + " / " + HeaderFooter.NumPages;

			//Zdefiniowanie stylów
			CellCustomStyle style = new CellCustomStyle(book);
			
			//Ustawienie szerokości kolumn oraz stylu
			for (int cw = 0; cw < columnsWidth.Count; cw++)
			{
				sheet.SetDefaultColumnStyle(cw, style.defaultStyle);
				sheet.SetColumnWidth(cw, (int)(columnsWidth[cw] * 42));
			}

			this.relativesCount = relativesCount;

			int length = 0;
			SectionADrawing(book, sheet, 0, style, sectionAHeights, contentSectionA);
			if (relativesCount != 0)
			{
				for (int r = 0; r < relativesCount; r++)
				{
					if (r == 0)
					{
						length += sectionAHeights.Count;
					}
					else
					{
						length += sectionBHeights.Count;
						sheet.SetRowBreak(length);
					}
					SectionBDrawing(book, sheet, length, style, sectionBHeights, contentSectionB, r + 2);
				}

				length += sectionBHeights.Count;
				//if (relativesCount % 2 != 0 && relativesCount != 3 && relativesCount != 5 &&
				//	relativesCount != 7 && relativesCount != 13 && relativesCount != 15 && relativesCount != 11 ||  
				//	relativesCount == 4 ||  relativesCount == 12 || relativesCount == 16 || relativesCount == 6 ||
				//	relativesCount == 14)
				//{
					sheet.SetRowBreak(length);
				//}
			}
			else
			{
				length += sectionAHeights.Count;
				SectionBDrawingAlt(book, sheet, length, style, sectionBHeightsAlt, contentSectionBAlt, 2);
				length += sectionBHeightsAlt.Count;
			}


			SectionCDrawing(book, sheet, length, style, sectionCHeights, contentSectionC);
			length += sectionCHeights.Count;
			

			book.SetPrintArea(0, 2, columnsWidth.Count-1, 0, length);

			return book;
		}

		public static IWorkbook CreateNewBook(string filePath)
		{
			IWorkbook book;
			var extension = Path.GetExtension(filePath);

			// HSSF => Microsoft Excel(xls)(excel 97-2003)
			// XSSF => Office Open XML Workbook(xlsx)(excel 2007)
			if (extension == ".xls") {
				book = new HSSFWorkbook();
			}
			else if (extension == ".xlsx") {
				book = new XSSFWorkbook();
			}
			else {
				throw new ApplicationException("CreateNewBook: invalid extension");
			}

			return book;
		}

		public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, string value)
		{
			var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
			var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

			cell.SetCellValue(value);
		}

		public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, double value)
		{
			var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
			var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

			cell.SetCellValue(value);
		}

		public static void WriteCell(ISheet sheet, int columnIndex, int rowIndex, DateTime value)
		{
			var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
			var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

			cell.SetCellValue(value);
		}

		public static void WriteStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
		{
			var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
			var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

			cell.CellStyle = style;
		}

		public static void SetRowHeight(ISheet sheet, int rowIndex, int height)
		{
			var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
			row.HeightInPoints = (short)height;
		}

		public static void MergeRegion(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
		{
			NPOI.SS.Util.CellRangeAddress region = new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
			sheet.AddMergedRegion(region);
		}

		public static void SectionCDrawing( IWorkbook book, ISheet sheet, int start, CellCustomStyle style, List<int> sectionCHeights, List<string> contentSectionC )
		{
			int rn = start; // Start Row number

			//Ustawienie wysokości wierszy sekcji C
			int rnTemp = rn;
			foreach (int v in sectionCHeights)
			{
				SetRowHeight(sheet, rnTemp, v);
				rnTemp++;
			}

			rn++;
			rn++;
			int mergeto = rn + 15;
			MergeRegion(sheet, rn, mergeto, 3, 3);
			for (int m = rn; m <= mergeto; m++)
			{
				WriteStyle(sheet, 3, m, style.titleRotatedAddStyle);
				if ( m == mergeto)
				{
					WriteStyle(sheet, 3, m, style.titleRotatedStyle);
				}
			}
			WriteCell(sheet, 5, rn, contentSectionC[0]);
			WriteStyle(sheet, 5, rn, style.titleStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[1]);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[2]);
			WriteStyle(sheet, 5, rn, style.statementAddStyle);
			WriteCell(sheet, 6, rn, contentSectionC[3]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[4]);
			WriteStyle(sheet, 5, rn, style.statementAddStyle);
			WriteCell(sheet, 6, rn, contentSectionC[5]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[6]);
			WriteStyle(sheet, 5, rn, style.statementAddStyle);
			WriteCell(sheet, 6, rn, contentSectionC[7]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 6, rn, contentSectionC[8]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 6, rn, contentSectionC[9]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[10]);
			WriteStyle(sheet, 5, rn, style.statementAddStyle);
			WriteCell(sheet, 6, rn, contentSectionC[11]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 6, rn, contentSectionC[12]);
			WriteStyle(sheet, 6, rn, style.statementStyle);

			rn++;
			WriteCell(sheet, 24, rn, contentSectionC[13]);

			rn++;
			SignatureLine( sheet, rn, style, 24, 33, "top");

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[14]);
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			SquaredLine(sheet, rn, style, 5, 22);
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[15]);
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			SquaredLine(sheet, rn, style, 5, 12, 6, 8);
			SignatureLine( sheet, rn, style, 24, 33, "bottom");

			rn++;
			rn++;
			rn++;
			SeparatorLine(sheet, rn, style, 4, 33);
			int mergetoo = rn + 7;
			MergeRegion(sheet, rn, mergetoo, 3, 3);
			for (int m = rn; m <= mergetoo; m++)
			{
				WriteStyle(sheet, 3, m, style.titleRotatedAddStyle);
				if ( m == mergetoo)
				{
					WriteStyle(sheet, 3, m, style.titleRotatedStyle);
				}
			}

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[16]);
			WriteStyle(sheet, 5, rn, style.titleStyle);

			rn++;
			WriteCell(sheet, 14, rn, contentSectionC[17]);
			WriteCell(sheet, 24, rn, contentSectionC[18]);

			rn++;
			SignatureLine( sheet, rn, style, 14, 22, "top");
			SignatureLine( sheet, rn, style, 24, 33, "top");

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[19]);
			SignatureLine( sheet, rn, style, 14, 22, "middle");
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			SquaredLine(sheet, rn, style, 5, 10);
			SignatureLine( sheet, rn, style, 14, 22, "middle");
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			WriteCell(sheet, 5, rn, contentSectionC[20]);
			SignatureLine( sheet, rn, style, 14, 22, "middle");
			SignatureLine( sheet, rn, style, 24, 33, "middle");

			rn++;
			SquaredLine(sheet, rn, style, 5, 12, 6, 8);
			SignatureLine( sheet, rn, style, 14, 22, "bottom");
			SignatureLine( sheet, rn, style, 24, 33, "bottom");
		}

		public static void SectionBDrawing(IWorkbook book, ISheet sheet, int start, CellCustomStyle style, List<int> sectionBHeights, List<string> contentSectionB, int secNum)
		{
			int rn = start; // Start Row number

			int rnTemp = rn;
			foreach (int v in sectionBHeights)
			{
				SetRowHeight(sheet, rnTemp, v);
				rnTemp++;
			}

			rn++;
			int mergeto = rn + 40;
			//Scalenie komórek tytułu sekcji B
			MergeRegion(sheet, rn, mergeto, 3, 3);
			for (int m = rn; m <= mergeto; m++)
			{
				WriteStyle(sheet, 3, m, style.titleRotatedStyle);
			}
			WriteCell(sheet, 3, rn, contentSectionB[0]);
			WriteCell(sheet, 5, rn, ToRoman(secNum) + contentSectionB[1]);
			WriteStyle(sheet, 5, rn, style.titleStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[2]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[3]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[4]);
			WriteCell(sheet, 17, rn, contentSectionB[5]);
			WriteCell(sheet, 26, rn, contentSectionB[6]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 15);
			SquaredLine(sheet, rn, style, 17, 24, 18, 20);
			SquaredLine(sheet, rn, style, 26, 28, 28, 28);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[7]);
			WriteCell(sheet, 15, rn, contentSectionB[8]);
			WriteCell(sheet, 24, rn, contentSectionB[9]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 13, 13, 13);
			SquaredLine(sheet, rn, style, 15, 22, 22, 22);
			WriteStyle(sheet, 24, rn, style.codeStyle);

			rn++;
			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[10]);
			WriteCell(sheet, 7, rn, contentSectionB[11]);

			for ( int o = 12; o <= 22; o++ ) {
				rn++;
				WriteStyle(sheet, 5, rn, style.codeStyle);
				WriteStyle(sheet, 6, rn, style.boxStyle);
				WriteStyle(sheet, 7, rn, style.optionStyle);
				WriteCell(sheet, 7, rn, contentSectionB[o]);
			}

			rn++;
			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[23]);
			WriteCell(sheet, 7, rn, contentSectionB[24]);

			for ( int o = 25; o <= 29; o++ ) {
				rn++;
				WriteStyle(sheet, 5, rn, style.codeStyle);
				WriteStyle(sheet, 6, rn, style.boxStyle);
				WriteStyle(sheet, 7, rn, style.optionStyle);
				WriteCell(sheet, 7, rn, contentSectionB[o]);
			}

			rn++;
			rn++;
			WriteCell(sheet, 5, rn, ToRoman(secNum) + contentSectionB[30]);
			WriteStyle(sheet, 5, rn, style.titleStyle);
			WriteCell(sheet, 12, rn, contentSectionB[31]);
			WriteStyle(sheet, 12, rn, style.noticeStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[32]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[33]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[34]);
			WriteCell(sheet, 28, rn, contentSectionB[35]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 26);
			SquaredLine(sheet, rn, style, 28, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[36]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionB[37]);
			WriteCell(sheet, 14, rn, contentSectionB[38]);
			WriteCell(sheet, 20, rn, contentSectionB[39]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 12);
			SquaredLine(sheet, rn, style, 14, 18);
			SquaredLine(sheet, rn, style, 20, 33);
		}

		public static void SectionBDrawingAlt(IWorkbook book, ISheet sheet, int start, CellCustomStyle style, List<int> sectionBHeightsAlt, List<string> contentSectionBAlt, int secNum)
		{
			int rn = start; // Start Row number

			int rnTemp = rn;
			foreach (int v in sectionBHeightsAlt)
			{
				SetRowHeight(sheet, rnTemp, v);
				rnTemp++;
			}

			rn++;
			int mergeto = rn + 3;
			//Scalenie komórek tytułu sekcji B
			MergeRegion(sheet, rn, mergeto, 3, 3);
			for (int m = rn; m <= mergeto; m++)
			{
				WriteStyle(sheet, 3, m, style.titleRotatedStyle);
			}
			WriteCell(sheet, 5, rn, ToRoman(secNum) + contentSectionBAlt[1]);
			WriteStyle(sheet, 5, rn, style.titleStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionBAlt[2]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			WriteContent( sheet, "NIE ZGŁASZAM NIKOGO".ToCharArray(), rn, 5);
		}

		public static void SectionADrawing(IWorkbook book, ISheet sheet, int start, CellCustomStyle style, List<int> sectionAHeights, List<string> contentSectionA)
		{
			int rn = start; // Start Row number

			//Ustawienie wysokości wierszy sekcji A
			int rnTemp = rn;
			foreach (int v in sectionAHeights)
			{
				SetRowHeight(sheet, rnTemp, v);
				rnTemp++;
			}

			//Rysowanie sekcji A
			WriteCell(sheet, 9, rn, contentSectionA[0]);
			WriteStyle(sheet, 9, rn, style.introductionStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionA[1]);
			SquaredLine(sheet, rn, style, 5, 7, true);

			WriteCell(sheet, 21, rn, contentSectionA[2]);
			SquaredLine(sheet, rn, style, 8, 33, true);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionA[3]);
			WriteCell(sheet, 29, rn, contentSectionA[4]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 12, 6, 8);

			SquaredLine(sheet, rn, style, 29, 33);

			rn++; //4 - pusty wiersz
			rn++; //5
			int mergeto = rn + 6;
			//Scalenie komórek tytułu sekcji A
			MergeRegion(sheet, rn, mergeto, 3, 3);
			for (int m = rn; m <= mergeto; m++)
			{
				WriteStyle(sheet, 3, m, style.titleRotatedStyle);
			}
			WriteCell(sheet, 3, rn, contentSectionA[5]);
			WriteCell(sheet, 5, rn, contentSectionA[6]);
			WriteStyle(sheet, 5, rn, style.titleStyle);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionA[7]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionA[8]);

			rn++; //9
			SquaredLine(sheet, rn, style, 5, 33);

			rn++;
			WriteCell(sheet, 5, rn, contentSectionA[9]);

			rn++;
			SquaredLine(sheet, rn, style, 5, 33);

		}

		public static void SquaredLine(ISheet sheet, int rn, CellCustomStyle style, int startCell, int endCell)
		{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell) {
						WriteStyle(sheet, m, rn, style.squaredLStyle);
					} else
					{
						WriteStyle(sheet, m, rn, style.squaredMStyle);
					}

					if (m == endCell) {
						WriteStyle(sheet, m, rn, style.squaredRStyle);
					}
				}
		}

		public static void SquaredLine(ISheet sheet, int rn, CellCustomStyle style, int startCell, int endCell, bool title)
		{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell) {
						WriteStyle(sheet, m, rn, style.titleLStyle);
					} else
					{
						WriteStyle(sheet, m, rn, style.titleMStyle);
					}

					if (m == endCell) {
						WriteStyle(sheet, m, rn, style.titleRStyle);
					}
				}
		}

		public static void SquaredLine(ISheet sheet, int rn, CellCustomStyle style, int startCell, int endCell, int separator, int separatorAdd)
		{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell) {
						WriteStyle(sheet, m, rn, style.squaredSpecialLStyle);
					} else
					{
						WriteStyle(sheet, m, rn, style.squaredSpecialMStyle);
					}

					if (m == endCell || m == separator || m == separatorAdd) {
						WriteStyle(sheet, m, rn, style.squaredSpecialRStyle);
					}
				}
		}

		public static void SignatureLine(ISheet sheet, int rn, CellCustomStyle style, int startCell, int endCell, string type)
		{
			if (type.Equals("top"))
			{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell)
					{
						WriteStyle(sheet, m, rn, style.signatureLTCornerStyle);
					}
					else
					{
						WriteStyle(sheet, m, rn, style.signatureTStyle);
					}

					if (m == endCell)
					{
						WriteStyle(sheet, m, rn, style.signatureRTCornerStyle);
					}
				}
			}
			if (type.Equals("middle"))
			{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell)
					{
						WriteStyle(sheet, m, rn, style.signatureLStyle);
					}

					if (m == endCell)
					{
						WriteStyle(sheet, m, rn, style.signatureRStyle);
					}
				}
			}
			if (type.Equals("bottom"))
			{
				for (int m = startCell; m <= endCell; m++)
				{
					if (m == startCell)
					{
						WriteStyle(sheet, m, rn, style.signatureLBCornerStyle);
					}
					else
					{
						WriteStyle(sheet, m, rn, style.signatureBStyle);
					}

					if (m == endCell)
					{
						WriteStyle(sheet, m, rn, style.signatureRBCornerStyle);
					}
				}
			}
		}

		public static void SeparatorLine(ISheet sheet, int rn, CellCustomStyle style, int startCell, int endCell)
		{
			for (int m = startCell; m <= endCell; m++)
			{
				if (m == startCell) {
					WriteStyle(sheet, m, rn, style.separatorStyle);
				} else
				{
					WriteStyle(sheet, m, rn, style.separatorStyle);
				}
							
				if (m == endCell) {
					WriteStyle(sheet, m, rn, style.separatorStyle);
				}
			}
		}

		public static string ToRoman(int number)
		{
			if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
			if (number < 1) return string.Empty;            
			if (number >= 1000) return "M" + ToRoman(number - 1000);
			if (number >= 900) return "CM" + ToRoman(number - 900); //EDIT: i've typed 400 instead 900
			if (number >= 500) return "D" + ToRoman(number - 500);
			if (number >= 400) return "CD" + ToRoman(number - 400);
			if (number >= 100) return "C" + ToRoman(number - 100);            
			if (number >= 90) return "XC" + ToRoman(number - 90);
			if (number >= 50) return "L" + ToRoman(number - 50);
			if (number >= 40) return "XL" + ToRoman(number - 40);
			if (number >= 10) return "X" + ToRoman(number - 10);
			if (number >= 9) return "IX" + ToRoman(number - 9);
			if (number >= 5) return "V" + ToRoman(number - 5);
			if (number >= 4) return "IV" + ToRoman(number - 4);
			if (number >= 1) return "I" + ToRoman(number - 1);
			throw new ArgumentOutOfRangeException("something bad happened");
		}
	}

	class CellCustomStyle
	{

		public ICellStyle defaultStyle;
		public ICellStyle introductionStyle;
		public ICellStyle titleStyle;
		public ICellStyle titleRotatedStyle;
		public ICellStyle titleRotatedAddStyle;
		public ICellStyle titleLStyle;
		public ICellStyle titleMStyle;
		public ICellStyle titleRStyle;
		public ICellStyle squaredLStyle;
		public ICellStyle squaredMStyle;
		public ICellStyle squaredRStyle;
		public ICellStyle squaredSpecialLStyle;
		public ICellStyle squaredSpecialMStyle;
		public ICellStyle squaredSpecialRStyle;
		public ICellStyle codeStyle;
		public ICellStyle boxStyle;
		public ICellStyle optionStyle;
		public ICellStyle noticeStyle;
		public ICellStyle statementStyle;
		public ICellStyle statementAddStyle;
		public ICellStyle signatureLTCornerStyle;
		public ICellStyle signatureRTCornerStyle;
		public ICellStyle signatureLBCornerStyle;
		public ICellStyle signatureRBCornerStyle;
		public ICellStyle signatureRStyle;
		public ICellStyle signatureLStyle;
		public ICellStyle signatureTStyle;
		public ICellStyle signatureBStyle;
		public ICellStyle separatorStyle;

		public CellCustomStyle(IWorkbook book)
		{
			//Zdefiniowanie stylów
			defaultStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
										VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Left);
			introductionStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index);
			titleStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index);
			titleRotatedStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Thick,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Thick,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, 90);
			titleRotatedAddStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Thick,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, 90);
			titleLStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			titleMStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			titleRStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredLStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Thin,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredMStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Thin,
															 NPOI.SS.UserModel.BorderStyle.Thin, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredRStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.Thin, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredSpecialLStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Dotted,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredSpecialMStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Dotted,
															 NPOI.SS.UserModel.BorderStyle.Dotted, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			squaredSpecialRStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.Dotted, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			codeStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.Tan.Index, NPOI.HSSF.Util.HSSFColor.Tan.Index,
										VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center);
			boxStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			optionStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
										VerticalAlignment.Bottom, NPOI.SS.UserModel.HorizontalAlignment.Left);
			noticeStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index);
			statementStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
										VerticalAlignment.Bottom, NPOI.SS.UserModel.HorizontalAlignment.Left);
			statementAddStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
										VerticalAlignment.Bottom, NPOI.SS.UserModel.HorizontalAlignment.Right);

			signatureLTCornerStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureRTCornerStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureLBCornerStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureRBCornerStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureRStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureLStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureTStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);
			signatureBStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Medium,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index,
															 NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, NPOI.HSSF.Util.HSSFColor.Grey40Percent.Index, 0);

			separatorStyle = CreateStyle(book, FillPattern.SolidForeground, NPOI.HSSF.Util.HSSFColor.White.Index, NPOI.HSSF.Util.HSSFColor.White.Index,
															 VerticalAlignment.Center, NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.BorderStyle.None,
															 NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle.Double, NPOI.SS.UserModel.BorderStyle.Double,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index,
															 NPOI.HSSF.Util.HSSFColor.Black.Index, NPOI.HSSF.Util.HSSFColor.Black.Index, 0);

			//Zdefiniowanie czcionek
			IFont defaultFont = CreateFont(book, 8, "Arial Narrow", FontBoldWeight.Normal);
			IFont introductionFont = CreateFont(book, 9, "Arial Unicode MS", FontBoldWeight.Normal);
			IFont titleRotatedFont = CreateFont(book, 11, "Arial Narrow", FontBoldWeight.Normal);
			IFont titleSectionFont = CreateFont(book, 11, "Arial Narrow", FontBoldWeight.Bold);
			IFont titleMSectionFont = CreateFont(book, 12, "Arial Narrow", FontBoldWeight.Bold);
			IFont contentFont = CreateFont(book, 14, "Arial Narrow", FontBoldWeight.Bold);
			IFont optionFont = CreateFont(book, 11, "Arial Narrow", FontBoldWeight.Normal);
			IFont statementFont = CreateFont(book, 11, "Arial Narrow", FontBoldWeight.Bold);
			IFont noticeFont = CreateFont(book, 11, "Arial Narrow", FontBoldWeight.Normal);

			//Ustawienie czcionki dla stylu
			defaultStyle.SetFont(defaultFont);
			introductionStyle.SetFont(introductionFont);
			titleStyle.SetFont(titleSectionFont);
			titleRotatedStyle.SetFont(titleRotatedFont);
			titleLStyle.SetFont(titleSectionFont);
			titleMStyle.SetFont(titleMSectionFont);
			titleRStyle.SetFont(titleSectionFont);
			squaredLStyle.SetFont(contentFont);
			squaredMStyle.SetFont(contentFont);
			squaredRStyle.SetFont(contentFont);
			squaredSpecialLStyle.SetFont(contentFont);
			squaredSpecialMStyle.SetFont(contentFont);
			squaredSpecialRStyle.SetFont(contentFont);
			codeStyle.SetFont(contentFont);
			boxStyle.SetFont(contentFont);
			optionStyle.SetFont(optionFont);
			noticeStyle.SetFont(noticeFont);
			statementStyle.SetFont(statementFont);
			statementAddStyle.SetFont(statementFont);
		}

		public static IFont CreateFont(IWorkbook book, int fontHeightInPoints, string fontName, FontBoldWeight weight)
		{
			IFont font = book.CreateFont();
			font.FontHeightInPoints = (short)fontHeightInPoints;
			font.FontName = fontName;
			font.Boldweight = (short)weight;
			return font;
		}

		public static ICellStyle CreateStyle(IWorkbook book, FillPattern pattern, short fillBackgroundColor, short fillForegroundColor)
		{
			ICellStyle style = book.CreateCellStyle();
			style.FillPattern = pattern;
			style.FillBackgroundColor = fillBackgroundColor;
			style.FillForegroundColor = fillForegroundColor;
			return style;
		}

		public static ICellStyle CreateStyle(IWorkbook book, FillPattern pattern, short fillBackgroundColor, short fillForegroundColor,
												VerticalAlignment verticalAlignment, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment)
		{
			ICellStyle style = book.CreateCellStyle();
			style.FillPattern = pattern;
			style.FillBackgroundColor = fillBackgroundColor;
			style.FillForegroundColor = fillForegroundColor;

			style.VerticalAlignment = verticalAlignment;
			style.Alignment = horizontalAlignment;
			return style;
		}

		public static ICellStyle CreateStyle(IWorkbook book, FillPattern pattern, short fillBackgroundColor, short fillForegroundColor,
												VerticalAlignment verticalAlignment, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment,
												NPOI.SS.UserModel.BorderStyle borderRight, NPOI.SS.UserModel.BorderStyle borderLeft,
												NPOI.SS.UserModel.BorderStyle borderTop, NPOI.SS.UserModel.BorderStyle borderBottom,
												short rightBorderColor, short leftBorderColor, short topBorderColor, short bottomBorderColor,
												int rotation)
		{
			ICellStyle style = book.CreateCellStyle();
			style.FillPattern = pattern;
			style.FillBackgroundColor = fillBackgroundColor;
			style.FillForegroundColor = fillForegroundColor;

			style.VerticalAlignment = verticalAlignment;
			style.Alignment = horizontalAlignment;

			style.BorderRight = borderRight;
			style.BorderLeft = borderLeft;
			style.BorderTop = borderTop;
			style.BorderBottom = borderBottom;

			style.RightBorderColor = rightBorderColor;
			style.LeftBorderColor = leftBorderColor;
			style.TopBorderColor = topBorderColor;
			style.BottomBorderColor = bottomBorderColor;
			style.Rotation = (short)rotation;
			return style;
		}
	}
}
