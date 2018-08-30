package gvl.dom.report_engine.reports;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import gvl.dom.report_engine.ReportExtractingTiedAgencyPerformanceSegmentReport;
import main.utils.MySQLConnect;
import main.utils.ResultSetToExcel;
import main.utils.XLSXReadWriteHelper;

/**
 * Hello world!
 *
 */
public class TiedAgencyPerformanceReportV2 {
	final static Logger logger = Logger.getLogger(TiedAgencyPerformanceReportV2.class);
//	String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_template.xls";
//	String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_2018-07-31-RESULT.xls";

	public void fetchDataForGASheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "6.0 GA Performance";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_ga_performance(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void updateCoverSheet(String excelTemplate, String excelReport, String inputPeriodTo) {
		final String SHEET_NAME = "Cover";
		final int asAtRowIdx = 6;
		final int asAtColIdx = 5;
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;

		// format date
		SimpleDateFormat sdfMMMyy = new SimpleDateFormat("yyyy-MM-dd");

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			// CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("dd-mm-yyyy"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// reading sheet retention data
			XSSFSheet sheet = book.getSheet(SHEET_NAME);
			XSSFCell cell = sheet.getRow(asAtRowIdx).getCell(asAtColIdx);
			cell.setCellValue(sdfMMMyy.parse(inputPeriodTo));
			cell.setCellStyle(cellStyle2);

			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			book.write(fos);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			// Close workbook, OutputStream and Excel file to prevent leak
			try {
				fos.close();
				book.close();
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataApeMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_ape_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataFypMomSheet(String excelTemplate, String excelReport, String sheetName,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_fyp_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataCasecountMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_casecount_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataNewrecruitMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_newrecruit_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataManpowerMomSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_manpower_mom";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_manpower_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataActiveRatioMomSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_group_active_ratio_mom";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataActiveRatioYoySheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_group_active_ratio_yoy";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_alllevels_ytd(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataActiveRatioSAexclMomSheet(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_active_ratio_sa_excl_mom";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_activeratio_saexcle_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataActiveRatioSAexclYoySheet(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String SHEET_NAME = "data_active_ratio_sa_excl_yoy";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_activeratio_sa_excluded_alllevels_ytd(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataCasesizeMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_casesize_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataCaseperactiveMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_caseperactive_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataApeperactiveMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_apeperactive_alllevels(\"{0}\", \"{1}\", {2});",
					inputPeriodFrom, inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataActiveMomSheet(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo, int rowindex) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_active_alllevels(\"{0}\", \"{1}\", {2});", inputPeriodFrom,
					inputPeriodTo, rowindex);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataMPbyDesignationSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "group_manpower_by_desc_monthly";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_dynamic_man_power_by_designation_alllevels(\"{0}\", \"{1}\" );", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataRecruitmentSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "recruitment_monthly";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_dynamic_recruitment_by_designation_alllevels(\"{0}\", \"{1}\" );", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataRecruitmentActiveALSheet(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "active_recruit_leader_monthly";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_dynamic_activealrecruit_alllevels(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForDataRookie90daysSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_rookie90days";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_dynamic_rookie_performance_tiedagency_alllevels(\"{0}\", \"{1}\" );", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForProductMixAPESheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_product_mix_ape";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_product_mix_group_level(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForProductMixRiderSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_product_rider_mix";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_product_mix_group_level_rider_v1(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForProductMixTotalSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_product_mix_total";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_product_mix_group_level_total(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForProductMixCountSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_product_rider_mix_count";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_product_mix_group_level_rider_v1_count_products(\"{0}\", \"{1}\" );", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForAPESegDtlSheet(String excelTemplate, String excelReport, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_in_segment_ape";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_ape_report_dynamic(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForMPSegDtlSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_in_segment_manpower";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_mp_report_dynamic(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForCasecountSegDtlSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_in_segment_cscnt";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_casecount_report_dynamic(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForActiveSegDtlSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_in_segment_active";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_active_report_dynamic(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_detail_seg_chart_y0_ape(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_seg_chart_y0_ape";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_ape(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_detail_seg_chart_y0_mp(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_seg_chart_y0_mp";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_mp(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_detail_seg_chart_y0_case(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_seg_chart_y0_case";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_case(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_detail_seg_chart_y0_act(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		final String sheetname = "data_detail_seg_chart_y0_act";

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_active(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_group_manpower_monthly_total(String excelTemplate, String excelReport,
			String sheetname, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_mp_tt(\"{0}\", \"{1}\" , 7);",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

	public void fetchDataForSheet_data_recruitment_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_recruitment_tt(\"{0}\", \"{1}\" );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	public void fetchDataForSheet_data_ape_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_tt(\"{0}\", \"{1}\", 3 );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	public void fetchDataForSheet_data_ar_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_activeratio(\"{0}\", \"{1}\", 3 );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	public void fetchDataForSheet_data_casesize_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_casesize(\"{0}\", \"{1}\", 3 );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	public void fetchDataForSheet_data_caseperactive_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_case_per_active(\"{0}\", \"{1}\", 3 );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	public void fetchDataForSheet_data_apeperactive_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;
		
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			// CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle2 = null;
			// CellStyle cellStyle3 = book.createCellStyle();
			// CreationHelper createHelper = book.getCreationHelper();
			// cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_per_active(\"{0}\", \"{1}\", 3 );",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
				mySQLConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}
	
	public synchronized void writeDataForSheet(String excelTemplate, String excelReport, String sheetname, int columnIdx, int rowIdx, ResultSet rs) {
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		
		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			CellStyle cellStyle2 = null;
			
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, rowIdx, columnIdx, rs);
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
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
	
	public void fetchDataForSheet_active_total(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_active_tt(\"{0}\", \"{1}\" , 7);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	
	}
	public void fetchDataForSheetPlanAPEtotal(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_ape(\"{0}\", 7);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	public void fetchDataForSheetPlanFYPtotal(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_fyp(\"{0}\", 8);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	public void fetchDataForSheetPlanCscnttotal(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_cscnt(\"{0}\", 9);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	public void fetchDataForSheetPlanNewAgenttotal(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_newagents(\"{0}\", 10);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	public void fetchDataForSheetPlanMPtotal(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_mp(\"{0}\", 11);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void fetchDataForSheetPlanAPEtotalYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_ape_ytd(\"{0}\", \"{1}\", 7);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanFYPtotalYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_fyp_ytd(\"{0}\", \"{1}\", 8);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanCSCNTtotalYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_cscnt_ytd(\"{0}\", \"{1}\", 9);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanNewAgentsttotalYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_newagents_ytd(\"{0}\", \"{1}\", 10);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void fetchDataForSheetPlanAR(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_ar(\"{0}\", 7);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanCasesize(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_casesize(\"{0}\", 9);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanCasePAct(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_caseperactive(\"{0}\", 10);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanApePAct(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_apeperactive(\"{0}\", 11);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanAct(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_active(\"{0}\", 12);",inputPeriodFrom);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void fetchDataForSheetPlanActiveYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_activeag_ytd(\"{0}\", \"{1}\", 12);",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanAPEYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_ape_fullyear(\"{0}\", \"{1}\");",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanMPYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_mp_fullyear(\"{0}\", \"{1}\");",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetPlanNewAgentsYTD(String excelTemplate, String excelReport, String sheetname,
			String inputPeriodFrom, String inputPeriodTo) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_target_newrecruit_fullyear(\"{0}\", \"{1}\");",inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetAdlist(String excelTemplate, String excelReport, String sheetname) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_getadlist();", "", "");
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetSalesChart(String excelTemplate, String excelReport, String sheetname) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_getsaleschart();", "", "");
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, 2, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void fetchDataForSheetUserReport(String excelTemplate, String excelReport, String sheetname) {
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for sheet: " + sheetname);
		}
		try {
			MySQLConnect mySQLConnect = null;
			String sqlcommand = null;
			ResultSet rs = null;
			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);
			// ending manpower
			sqlcommand = MessageFormat.format("call tiedagency_dynamic_getuserreport();", "", "");
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			writeDataForSheet(excelReport, excelReport, sheetname, -1, 0, rs);
			mySQLConnect.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
