/**
 * 
One interesting thing with POI is that (I don't know whether it's intentional, accidentally or real) 
it has got some really funny names for its workbook implementations e.g.

XSSF (XML SpreadSheet Format) – Used to reading and writting Open Office XML (XLSX) format files.    
HSSF (Horrible SpreadSheet Format) – Use to read and write Microsoft Excel (XLS) format files.
HWPF (Horrible Word Processor Format) – to read and write Microsoft Word 97 (DOC) format files.
HSMF (Horrible Stupid Mail Format) – pure Java implementation for Microsoft Outlook MSG files
HDGF (Horrible DiaGram Format) – One of the first pure Java implementation for Microsoft Visio binary files.    
HPSF (Horrible Property Set Format) – For reading “Document Summary” information from Microsoft Office files. 
HSLF (Horrible Slide Layout Format) – a pure Java implementation for Microsoft PowerPoint files.
HPBF (Horrible PuBlisher Format) – Apache's pure Java implementation for Microsoft Publisher files.
DDF (Dreadful Drawing Format) – Apache POI package for decoding the Microsoft Office Drawing format.

It's very important that you know full form of these acronyms, otherwise it would be difficult to keep track of 
which implementation is for which format. If you are only concerned about reading Excel files 
then at-least remember XSSF and HSSF classes e.g. XSSFWorkBook and HSSFWorkBook.

Read more: http://www.java67.com/2014/09/how-to-read-write-xlsx-file-in-java-apache-poi-example.html#ixzz5AHCqrJVI
 */
package main.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.Date;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** * Java program to read and write Excel file in Java using Apache POI * */
public class XLSXReadWriteHelper {

	/**
	 * Create a new excel workbook contains sheets in the sheetNames
	 * 
	 * @param excelFileName
	 * @param sheetNames
	 * @throws IOException
	 */
	public void createXLSX(String excelFileName, String[] sheetNames) throws IOException {
		// create a new workbook
		XSSFWorkbook book = new XSSFWorkbook();
		for (String sheetName : sheetNames) {
			// create sheets in the workbook
			book.createSheet(sheetName);
		}
		// open an OutputStream to save written data into Excel file
		File excel = new File(excelFileName);
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		// Close workbook, OutputStream and Excel file to prevent leak
		os.close();
		book.close();
	}

	
	public static ArrayList<Row> read(String excelFileName, String sheetName, int skipRows) throws IOException {
		ArrayList<Row> rows = new ArrayList<Row>();
		File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet(sheetName);
		Iterator<Row> itr = sheet.iterator();
		// Iterating over Excel file in Java
		while (itr.hasNext()) {
			Row row = itr.next();
			if (row.getRowNum() < skipRows) {
				continue;
			}
			rows.add(row);
		}
		// Close workbook, OutputStream and Excel file to prevent leak
		book.close();
		fis.close();
		return rows;
	}

	public Map<String, Object> read(File excel, String sheetNames[], int skipRows[]) throws IOException {
		Map<String, Object> map = new HashMap<String, Object>();
		// File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		for (int i = 0; i < sheetNames.length; i++) {
			ArrayList<Row> rows = new ArrayList<Row>();
			String sheetName = sheetNames[i];
			int skipRow = skipRows[i];
			XSSFSheet sheet = book.getSheet(sheetName);
			Iterator<Row> itr = sheet.iterator();
			// Iterating over Excel file in Java
			while (itr.hasNext()) {
				Row row = itr.next();
				if (row.getRowNum() < skipRow) {
					continue;
				}
				rows.add(row);
			}
			map.put(sheetName, rows);
		}
		// Close workbook, OutputStream and Excel file to prevent leak
		book.close();
		fis.close();
		return map;
	}

	/**
	 * Read data from multi sheets
	 * 
	 * @param excel
	 * @param sheetNames
	 * @param skipRows:
	 *            number of rows skipped reading
	 * @param maxRows:
	 *            number of read rows, -1 to read all rows from the skip row to the
	 *            end
	 * @return
	 * @throws IOException
	 */
	public Map<String, Object> read(File excel, String sheetNames[], int skipRows[], int[] maxRows) throws IOException {
		Map<String, Object> map = new HashMap<String, Object>();
		// File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		for (int i = 0; i < sheetNames.length; i++) {
			ArrayList<Row> rows = new ArrayList<Row>();
			String sheetName = sheetNames[i];
			int skipRow = skipRows[i];
			int maxRow = maxRows[i];
			XSSFSheet sheet = book.getSheet(sheetName);
			Iterator<Row> itr = sheet.iterator();
			// Iterating over Excel file in Java
			while (itr.hasNext()) {
				Row row = itr.next();
				if (row.getRowNum() < skipRow) {
					continue;
				}
				rows.add(row);
				if (rows.size() == maxRow) {
					break;
				}
			}
			map.put(sheetName, rows);
		}
		// Close workbook, OutputStream and Excel file to prevent leak
		book.close();
		fis.close();
		return map;
	}

	public ArrayList<Row> read(File excel, String sheetName, int skipRows, int maxRows) throws IOException {
		ArrayList<Row> rows = new ArrayList<Row>();
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet(sheetName);
		Iterator<Row> itr = sheet.iterator();
		// Iterating over Excel file in Java
		while (itr.hasNext()) {
			Row row = itr.next();
			if (row.getRowNum() < skipRows) {
				continue;
			}
			rows.add(row);
			if (rows.size() == maxRows) {
				break;
			}
		}
		// Close workbook, OutputStream and Excel file to prevent leak
		book.close();
		fis.close();
		return rows;
	}

	public ArrayList<Row> read(File excel, String sheetName, int skipRows) throws IOException {
		ArrayList<Row> rows = new ArrayList<Row>();
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet(sheetName);
		Iterator<Row> itr = sheet.iterator();
		// Iterating over Excel file in Java
		while (itr.hasNext()) {
			Row row = itr.next();
			if (row.getRowNum() < skipRows) {
				continue;
			}
			rows.add(row);
		}
		// Close workbook, OutputStream and Excel file to prevent leak
		book.close();
		fis.close();
		return rows;
	}

	/**
	 * write array list of rows to existing excel file
	 * 
	 * @param excelFileName
	 * @param sheetName
	 * @param rownum
	 * @param rows
	 * @return
	 * @throws IOException
	 */
	public int write(String excelFileName, String sheetName, int rownum, ArrayList<Row> rows) throws IOException {
		File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet(sheetName);
		// 1. Create the date cell style
		CellStyle cellStyle = book.createCellStyle();
		CreationHelper createHelper = book.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));

		// writing data into XLSX file
		for (int i = 0; i < rows.size(); i++) {
			Row row = rows.get(i);
			Row newRow = sheet.getRow(rownum + i);
			if (newRow == null) {
				newRow = sheet.createRow(rownum + i);
			}
			// Iterating over each column of Excel file
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				Cell newCell = newRow.getCell(cell.getColumnIndex());
				if (newCell == null) {
					newCell = newRow.createCell(cell.getColumnIndex());
				}
				CellType cellTypeEnum = cell.getCellTypeEnum();
				switch (cellTypeEnum) {
				case STRING:
					newCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						newCell.setCellValue(cell.getDateCellValue());
						newCell.setCellStyle(cellStyle);
					} else {
						CellStyle cs = book.createCellStyle();
						CreationHelper ch = book.getCreationHelper();
						cs.setDataFormat(
								createHelper.createDataFormat().getFormat(cell.getCellStyle().getDataFormatString()));
						newCell.setCellValue(cell.getNumericCellValue());
						newCell.setCellStyle(cs);
					}
					break;
				case BOOLEAN:
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
				case FORMULA:
					switch (cell.getCachedFormulaResultTypeEnum()) {
					case NUMERIC:
						newCell.setCellValue(cell.getNumericCellValue());
						break;
					case STRING:
						newCell.setCellValue(cell.getRichStringCellValue());
						break;
					default:
						break;
					}
					break;
				default:
					break;
				}
			}
		}
		// open an OutputStream to save written data into Excel file
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		// Close workbook, OutputStream and Excel file to prevent leak
		os.close();
		book.close();
		fis.close();
		return rownum + rows.size();
	}

	/**
	 * Write [rows] to the [sheetName] of the [book] at the row number [rownum]
	 * 
	 * @param book
	 * @param cellStyle
	 * @param sheetName
	 * @param rownum
	 * @param rows
	 * @return
	 * @throws IOException
	 */
	public static int write(XSSFWorkbook book, CellStyle cellStyle, String sheetName, int rownum, ArrayList<Row> rows)
			throws IOException {
		XSSFSheet sheet = book.getSheet(sheetName);
		// 1. Create the date cell style

		CreationHelper createHelper = book.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
		CellStyle cs = book.createCellStyle();
		// writing data into XLSX file
		for (int i = 0; i < rows.size(); i++) {
			Row row = rows.get(i);
			// get the excel row from the sheet, if not exist then create one
			Row newRow = sheet.getRow(rownum + i);
			if (newRow == null) {
				newRow = sheet.createRow(rownum + i);
			}
			// Iterating over each column of Excel file
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				Cell newCell = newRow.getCell(cell.getColumnIndex());
				if (newCell == null) {
					newCell = newRow.createCell(cell.getColumnIndex());
				}
				CellType cellTypeEnum = cell.getCellTypeEnum();
				switch (cellTypeEnum) {
				case STRING:
					newCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						newCell.setCellValue(cell.getDateCellValue());
						newCell.setCellStyle(cellStyle);
					} else {
						// CellStyle cs = book.createCellStyle();
						cs.setDataFormat(
								createHelper.createDataFormat().getFormat(cell.getCellStyle().getDataFormatString()));
						newCell.setCellValue(cell.getNumericCellValue());
						newCell.setCellStyle(cs);
					}
					break;
				case BOOLEAN:
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
				case FORMULA:
					switch (cell.getCachedFormulaResultTypeEnum()) {
					case NUMERIC:
						newCell.setCellValue(cell.getNumericCellValue());
						break;
					case STRING:
						newCell.setCellValue(cell.getRichStringCellValue());
						break;
					default:
						break;
					}
					break;
				default:
					break;
				}
			}
		}
		return rownum + rows.size();
	}

	/**
	 * Write [rows] to the [sheetName] of the [book] at the row number [rownum]
	 * 
	 * @param book
	 * @param cellStyle
	 * @param sheetName
	 * @param rownum
	 * @param rows
	 * @return
	 * @throws IOException
	 * @throws SQLException 
	 */
	public static String write(XSSFWorkbook book, CellStyle cellStyle, String sheetName, int rownum, int colnum, ResultSet rows)
			throws IOException, SQLException {
		XSSFSheet sheet = book.getSheet(sheetName);
		// Refreshing all the formulas in a sheet
		sheet.setForceFormulaRecalculation(true);
		// get total column count from the result set rows
		int numberOfcols = rows.getMetaData().getColumnCount();
		// loop through each row
		while (rows.next()) {
			// check if result set is at the first row, then I write the column name first
			if (rows.isFirst()) {
				Row headerRow = sheet.getRow(rownum);
				if (headerRow == null) {
					headerRow = sheet.createRow(rownum);
				}
				for (int i = 1; i <= numberOfcols; ++i) {
					// each element will be set to a new cell, so I create a new cell/or get
					// available cell from the newRow
					Cell newCell = headerRow.getCell(colnum + i);
					if (newCell == null) {
						newCell = headerRow.createCell(colnum + i);
					}
					newCell.setCellValue(rows.getMetaData().getColumnName(i));
				}
				++rownum;
			}

			// get the excel row at the next position [rownum] from the sheet, if not exist
			// then create one
			Row newRow = sheet.getRow(rownum);
			if (newRow == null) {
				newRow = sheet.createRow(rownum);
			}
			// loop through each element of the row in the result set
			for (int i = 1; i <= numberOfcols; ++i) {
				Object element = rows.getObject(i);
				// each element will be set to a new cell, so I create a new cell/or get
				// available cell from the newRow
				Cell newCell = newRow.getCell(colnum + i);
				if (newCell == null) {
					newCell = newRow.createCell(colnum + i);
				}
				// check cell's value type and set to the new cell
				if (element instanceof String) {
					newCell.setCellValue((String) element);
					continue; //for
				} else if (element instanceof Integer) {
					newCell.setCellValue((Integer) element);
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if (element instanceof BigDecimal) {
					newCell.setCellValue(((BigDecimal) element).doubleValue());
					if (null != cellStyle) newCell.setCellStyle(cellStyle); 
					continue; //for
				} else if (element instanceof Double) {
					newCell.setCellValue(((Double) element));
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if (element instanceof Long) {
					Long l = (Long) element;
					newCell.setCellValue(l.doubleValue());
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if  (element instanceof Date) {
					newCell.setCellValue((Date) element);
					continue; //for
				}
			}
			++rownum;
		}
		
		return rownum + ":" + colnum;
	}
	/**
	 * Write [rows] to the [sheetName] of the [book] at the row number [rownum]
	 * 
	 * @param book
	 * @param cellStyle
	 * @param sheetName
	 * @param rownum
	 * @param rows
	 * @return
	 * @throws IOException
	 * @throws SQLException 
	 */
	public static String writeLargeFile(SXSSFWorkbook book, CellStyle cellStyle, String sheetName, int rownum, int colnum, ResultSet rows)
			throws IOException, SQLException {
		SXSSFSheet sheet = book.getSheet(sheetName);
//		SXSSFSheet sheet = book.createSheet(sheetName);
		// Refreshing all the formulas in a sheet
		sheet.setForceFormulaRecalculation(true);
		// get total column count from the result set rows
		int numberOfcols = rows.getMetaData().getColumnCount();
		System.out.println(numberOfcols);
		CreationHelper createHelper = book.getCreationHelper();
		CellStyle datestyle = book.createCellStyle(); 
		datestyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
		
		
		// loop through each row
		while (rows.next()) {
			// check if result set is at the first row, then I write the column name first
			if (rows.isFirst()) {
				Row headerRow = sheet.getRow(rownum);
				if (headerRow == null) {
					headerRow = sheet.createRow(rownum);
				}
				for (int i = 1; i <= numberOfcols; ++i) {
					// each element will be set to a new cell, so I create a new cell/or get
					// available cell from the newRow
					Cell newCell = headerRow.getCell(colnum + i);
					if (newCell == null) {
						newCell = headerRow.createCell(colnum + i);
					}
					newCell.setCellValue(rows.getMetaData().getColumnName(i));
				}
				++rownum;
			}
			
			// get the excel row at the next position [rownum] from the sheet, if not exist
			// then create one
			Row newRow = sheet.getRow(rownum);
			if (newRow == null) {
				newRow = sheet.createRow(rownum);
			}
			// loop through each element of the row in the result set
			for (int i = 1; i <= numberOfcols; ++i) {
				Object element = rows.getObject(i);
				// each element will be set to a new cell, so I create a new cell/or get
				// available cell from the newRow
				Cell newCell = newRow.getCell(colnum + i);
				if (newCell == null) {
					newCell = newRow.createCell(colnum + i);
				}
				// check cell's value type and set to the new cell
				if (element instanceof String) {
					newCell.setCellValue((String) element);
					continue; //for
				} else if (element instanceof Integer) {
					newCell.setCellValue((Integer) element);
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if (element instanceof BigDecimal) {
					newCell.setCellValue(((BigDecimal) element).doubleValue());
					if (null != cellStyle) newCell.setCellStyle(cellStyle); 
					continue; //for
				} else if (element instanceof Double) {
					newCell.setCellValue(((Double) element));
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if (element instanceof Long) {
					Long l = (Long) element;
					newCell.setCellValue(l.doubleValue());
					if (null != cellStyle) newCell.setCellStyle(cellStyle);
					continue; //for
				} else if  (element instanceof Date) {
					newCell.setCellValue((Date) element);
					newCell.setCellStyle(datestyle);
					continue; //for
				}
			}
			++rownum;
		}
		
		return rownum + ":" + colnum;
	}
	
	public int writeColumn(String excelFileName, String sheetName, int rownum, int colnum, Object[] columnData)
			throws IOException {
		File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheet(sheetName);
		for (int i = 0; i < columnData.length; i++) {
			Row row = sheet.getRow(rownum + i);
			if (row == null) {
				row = sheet.createRow(rownum + i);
			}
			Object obj = columnData[i];
			Cell cell = row.getCell(colnum);
			if (cell == null) {
				cell = row.createCell(colnum);
			}
			if (obj instanceof String) {
				cell.setCellValue((String) obj);
			} else if (obj instanceof Boolean) {
				cell.setCellValue((Boolean) obj);
			} else if (obj instanceof Date) {
				cell.setCellValue((Date) obj);
			} else if (obj instanceof Double) {
				cell.setCellValue((Double) obj);
			}
		}
		// open an OutputStream to save written data into Excel file
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		// Close workbook, OutputStream and Excel file to prevent leak
		os.close();
		book.close();
		fis.close();
		return rownum + columnData.length;
	}

	/**
	 * write array list of rows to new excel file
	 * 
	 * @param excelFileName
	 * @param sheetName
	 * @param rownum
	 * @param rows
	 * @return
	 * @throws IOException
	 */
	public int write_new(String excelFileName, String sheetName, int rownum, ArrayList<Row> rows) throws IOException {
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet(sheetName);

		// 1. Create the date cell style
		CellStyle cellStyle = book.createCellStyle();
		CreationHelper createHelper = book.getCreationHelper();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));

		// writing data into XLSX file
		for (Row row : rows) {
			Row newRow = sheet.createRow(rownum++);
			// Iterating over each column of Excel file
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				Cell newCell = newRow.createCell(cell.getColumnIndex());
				CellType cellTypeEnum = cell.getCellTypeEnum();
				switch (cellTypeEnum) {
				case STRING:
					newCell.setCellValue(cell.getStringCellValue());
					break;
				case NUMERIC:
					if (HSSFDateUtil.isCellDateFormatted(cell)) {
						newCell.setCellValue(cell.getDateCellValue());
						newCell.setCellStyle(cellStyle);
					} else {
						CellStyle cs = book.createCellStyle();
						CreationHelper ch = book.getCreationHelper();
						cs.setDataFormat(
								createHelper.createDataFormat().getFormat(cell.getCellStyle().getDataFormatString()));
						newCell.setCellValue(cell.getNumericCellValue());
						newCell.setCellStyle(cs);
					}
					break;
				case BOOLEAN:
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
				default:
				}
			}
		}
		// open an OutputStream to save written data into Excel file
		File excel = new File(excelFileName);
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		// Close workbook, OutputStream and Excel file to prevent leak
		os.close();
		book.close();
		return rownum;
	}

	/**
	 * Filter rows whose value of the cells at column [colIndex] equals (ignore
	 * case) to [filterVal]
	 * 
	 * @param rows
	 * @param filterVal
	 * @param colIndex
	 * @return
	 */
	public ArrayList<Row> filter(ArrayList<Row> rows, String filterVal, int colIndex) {
		ArrayList<Row> filteredRows = new ArrayList<Row>();
		for (Row row : rows) {
			String teamColValue = row.getCell(colIndex) == null ? "" : row.getCell(colIndex).getStringCellValue();
			if (filterVal.equalsIgnoreCase(teamColValue)) {
				filteredRows.add(row);
			}
		}
		return filteredRows;
	}

	public void resizeCol(String excelFileName, String[] sheets) throws IOException {
		File excel = new File(excelFileName);
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		for (String sheetName : sheets) {
			XSSFSheet sheet = book.getSheet(sheetName);
			XSSFRow row = sheet.getRow(sheet.getFirstRowNum());
			// Resize all columns to fit the content size
			for (int i = 0; i < row.getLastCellNum(); i++) {
				sheet.autoSizeColumn(i);
			}
		}
		// open an OutputStream to save written data into Excel file
		FileOutputStream os = new FileOutputStream(excel);
		book.write(os);
		os.close();
		book.close();
		fis.close();
	}

	/**
	 * Clone the xlsx excel file
	 */
	private static void copyXlsxFile() {
		String excelTemplateTiedAgencySegReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31.xlsx";
		String excelTiedAgencySegReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31-RESULT.xlsx";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		try {
			// open the template
			File fileTemplate = new File(excelTemplateTiedAgencySegReport);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);

			// write to the new file
			File fileSavedTo = new File(excelTiedAgencySegReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}
}
