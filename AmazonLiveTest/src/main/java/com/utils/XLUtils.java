package com.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtils {

	public static FileInputStream fi;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell cell;
	public static FileOutputStream fo;
	public static CellStyle style;

	public static int getNoOfRows(String xlfile, String xlsheet) throws IOException {
		int realrowcount, normalrowcount = 0;
		try {

			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);
			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return 0;
			} else {
				ws = wb.getSheetAt(sheetindex);
				realrowcount = ws.getLastRowNum();
				/*
				 * Making RowNo to Normal counting purpose and use bcz Rowno starts from 0 and
				 * making it 1 here
				 */
				normalrowcount = realrowcount + 1;
				return normalrowcount;
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			/* To stop Resource Leakage */
			wb.close();
			fi.close();
		}
		/*
		 * If catch Block executed then this return statement will come into play else
		 * alwz return from try block
		 */
		return 0;
	}

	public static int getNoOfCells(String xlfile, String xlsheet, int rownum) throws IOException {
		int cellcount = 0;
		int normalsayrowno;
		try {
			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);

			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return 0;
			} else {
				ws = wb.getSheetAt(sheetindex);
				/*
				 * Making RowNo to Normal counting purpose and use bcz rowno starts from 0 and
				 * making it 1 here
				 */
				normalsayrowno = rownum - 1;
				row = ws.getRow(normalsayrowno);
				cellcount = row.getLastCellNum();
				return cellcount;
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			/* To stop Resource Leakage */
			wb.close();
			fi.close();
		}
		/*
		 * If catch Block executed then this return statement will come into play else
		 * alwz return from try block
		 */
		return 0;
	}

	public static String getCellData(String xlfile, String xlsheet, int rownum, int colnum) throws IOException {

		String celldata = "";
		try {
			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);

			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return "";
			} else {
				ws = wb.getSheetAt(sheetindex);
				row = ws.getRow(rownum - 1);

				cell = row.getCell(colnum - 1);
				if (cell.getCellTypeEnum() == CellType.NUMERIC) {
					System.out.println("CellType.NUMERIC Executed");
					// System.out.println("In CellType.NUMERIC datatype= "+cell.getCellTypeEnum());
					double celldata1 = cell.getNumericCellValue();
					/* TypeCasting double data to int type */
					int cd1 = (int) celldata1;
					/* Converting int to String type */
					celldata = String.valueOf(cd1);
					/*
					 * Converting double to string format ..if i will convert using this i will get
					 * decimal point values too celldata = String.valueOf(celldata1);
					 */
				} else if (cell.getCellTypeEnum() == CellType.STRING) {

					System.out.println("CellType.STRING Executed");
					// System.out.println("In CellType.STRING datatype= "+cell.getCellTypeEnum());
					celldata = cell.getStringCellValue();
				} else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
					System.out.println("CellType.BOOLEAN Executed");
					/* Getting boolean cell value and converting it to String type */
					celldata = String.valueOf(cell.getBooleanCellValue());
				} else if (cell.getCellTypeEnum() == CellType.BLANK) {
					System.out.println("CellType.BLANK Executed");
					/* For Blank_Cell Manuaaly Sending NoData */
					celldata = "NoData";
				} else if (cell.getCellTypeEnum() == CellType.ERROR) {
					System.out.println("CellType.ERROR Executed");
					byte celldata3 = cell.getErrorCellValue();
					/* Converting Byte to String */
					celldata = Byte.toString(celldata3);
				} else if (cell.getCellTypeEnum() == CellType.FORMULA) {
					System.out.println("CellType.FORMULA Executed");
					celldata = cell.getCellFormula();
				} else if (cell.getCellTypeEnum() == CellType._NONE) {
					System.out.println("CellType._NONE Executed");
					celldata = "None";
				}
				return celldata;
			}

		} catch (Exception e) {
			// celldata = "CellNotInUse";
			e.printStackTrace();
			// System.out.println("If Nothing Matched Cell Type= "+cell.getCellTypeEnum());
		} finally {
			/* To stop Resource Leakage */
			wb.close();
			fi.close();
		}
		/*
		 * If catch Block executed then this return statement will come into play else
		 * alwz return from try block
		 */
		return celldata;
	}

	public static void setCellData(String xlfile, String xlsheet, int rownum, int colnum, String celldata)
			throws IOException {
		try {
			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);

			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return "";
			} else {
				ws = wb.getSheetAt(sheetindex);
				row=ws.createRow(rownum-1);
				row = ws.getRow(rownum-1);

				cell = row.createCell(colnum-1);
				cell.setCellValue(celldata);

				fo = new FileOutputStream(xlfile);
				wb.write(fo);
			
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			wb.close();
			fi.close();
			fo.close();
		}

	}

	public static void FillGreenColor(String xlfile, String xlsheet, int rownum, int colnum) throws IOException {
		try {
			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);

			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return "";
			} else {
				ws = wb.getSheetAt(sheetindex);
				/*
				 * Both subtracted by -1 bcz in Excel Actual Count starts from 0 index ,so for
				 * starting it from 1 for normal counting we made changes
				 */
				row = ws.getRow(rownum - 1);
				cell = row.getCell(colnum - 1);

				style = wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				cell.setCellStyle(style);

				fo = new FileOutputStream(xlfile);
				wb.write(fo);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			wb.close();
			fi.close();
			fo.close();
		}

	}
	
	public static void FillRedColor(String xlfile, String xlsheet, int rownum, int colnum) throws IOException {
		try {
			fi = new FileInputStream(xlfile);
			wb = new XSSFWorkbook(fi);

			int sheetindex = wb.getSheetIndex(xlsheet);

			if (sheetindex == -1) {
				// System.out.println("Sheet U r Providing is not Present in WorkBook");
				throw new SheetNameNotFoundException(
						"'" + xlsheet + "'" + " Sheet is not Present in WorkBook that You Provided");
				// return "";
			} else {
				ws = wb.getSheetAt(sheetindex);
				/*
				 * Both subtracted by -1 bcz in Excel Actual Count starts from 0 index ,so for
				 * starting it from 1 for normal counting we made changes
				 */
				row = ws.getRow(rownum - 1);
				cell = row.getCell(colnum - 1);

				style = wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				// style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				cell.setCellStyle(style);

				fo = new FileOutputStream(xlfile);
				wb.write(fo);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			wb.close();
			fi.close();
			fo.close();
		}

	}
}
