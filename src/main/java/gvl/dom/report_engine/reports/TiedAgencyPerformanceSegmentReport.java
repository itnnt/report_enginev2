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
public class TiedAgencyPerformanceSegmentReport {
	final static Logger logger = Logger.getLogger(TiedAgencyPerformanceSegmentReport.class);
//	String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_template.xlsx";
//	String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31-RESULT.xlsx";

	public void fetchDataForCountrySheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_APE_ROWINDEX = 47;
		final int SECTOR_FYP_ROWINDEX = 195;
		final int SECTOR_RYP_ROWINDEX = 183;
		final int SECTOR_APE_CONTRIBUTION_ROWINDEX = 61;
		final int SECTOR_MANPOWER_ROWINDEX = 74;
		final int SECTOR_ACTIVE_ROWINDEX = 86;
		final int SECTOR_CASE_ROWINDEX = 110;
		final int SECTOR_CASESIZE_ROWINDEX = 122;
		final int SECTOR_CASEPERACTIVE_ROWINDEX = 134;
		final int SECTOR_APEPERACTIVE_ROWINDEX = 158;
		final int SECTOR_APEPERMANPOWER_ROWINDEX = 146;
		final int SECTOR_ACTIVERATIO_ROWINDEX = 98;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "Country";

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
			CellStyle cellStylePercentage = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			cellStylePercentage.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format("call report_segmentation_man_power_by_designation(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment
			sqlcommand = MessageFormat.format(
					"call report_segmentation_recruitment_by_designation_tiedagency(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);

			// AL recruitment KPIs
			sqlcommand = MessageFormat.format(
					"call report_segmentation_al_recruitment_kpis_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// APE
			sqlcommand = MessageFormat.format(
					"call report_ape_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// fyp
			sqlcommand = MessageFormat.format("call report_fyp_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_FYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// ryp
			sqlcommand = MessageFormat.format("call report_ryp_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// APE contribution
			sqlcommand = MessageFormat.format(
					"call report_ape_contribution_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_APE_CONTRIBUTION_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// manpower
			sqlcommand = MessageFormat.format(
					"call report_manpower_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_MANPOWER_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// active
			sqlcommand = MessageFormat.format(
					"call report_active_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ACTIVE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// case
			sqlcommand = MessageFormat.format(
					"call report_case_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// casesize
			sqlcommand = MessageFormat.format(
					"call report_casesize_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASESIZE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// case per active
			sqlcommand = MessageFormat.format(
					"call report_caseperactive_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASEPERACTIVE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// ape per active
			sqlcommand = MessageFormat.format(
					"call report_apeperactive_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERACTIVE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// ape per manpower
			sqlcommand = MessageFormat.format(
					"call report_apepermanpower_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERMANPOWER_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// active ratio
			sqlcommand = MessageFormat.format("call report_activeratio_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_ACTIVERATIO_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
						
			// Rookie Performance
			sqlcommand = MessageFormat.format(
					"call report_segmentation_rookie_performance_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

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

	public void fetchDataForNorthSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_APE_ROWINDEX = 47;
		final int SECTOR_FYP_ROWINDEX = 195;
		final int SECTOR_RYP_ROWINDEX = 183;
		final int SECTOR_APE_CONTRIBUTION_ROWINDEX = 61;
		final int SECTOR_MANPOWER_ROWINDEX = 74;
		final int SECTOR_ACTIVE_ROWINDEX = 86;
		final int SECTOR_CASE_ROWINDEX = 110;
		final int SECTOR_CASESIZE_ROWINDEX = 122;
		final int SECTOR_CASEPERACTIVE_ROWINDEX = 134;
		final int SECTOR_APEPERMANPOWER_ROWINDEX = 146;
		final int SECTOR_APEPERACTIVE_ROWINDEX = 158;
		final int SECTOR_ACTIVERATIO_ROWINDEX = 98;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "North";

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
			CellStyle cellStylePercentage = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			cellStylePercentage.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_segmentation_man_power_by_designation_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment
			sqlcommand = MessageFormat.format(
					"call report_segmentation_recruitment_by_designation_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);

			// AL recruitment KPIs
			sqlcommand = MessageFormat.format(
					"call report_segmentation_al_recruitment_kpis_tiedagency_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// APE
			sqlcommand = MessageFormat.format("call report_ape_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APE_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// fyp
			sqlcommand = MessageFormat.format("call report_fyp_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_FYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// ryp
			sqlcommand = MessageFormat.format("call report_ryp_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// APE contribution
			sqlcommand = MessageFormat.format("call report_ape_contribution_north(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_APE_CONTRIBUTION_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

			// manpower
			sqlcommand = MessageFormat.format("call report_manpower_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_MANPOWER_ROWINDEX, SECTOR_COLUMNINDEX, rs);
						
			// active
			sqlcommand = MessageFormat.format("call report_active_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// case
			sqlcommand = MessageFormat.format("call report_case_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// casesize
			sqlcommand = MessageFormat.format("call report_casesize_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASESIZE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// case per active
			sqlcommand = MessageFormat.format("call report_caseperactive_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASEPERACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// APE per active
			sqlcommand = MessageFormat.format("call report_apeperactive_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// APE per manpower 
			sqlcommand = MessageFormat.format("call report_apepermanpower_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERMANPOWER_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// active ratio
			sqlcommand = MessageFormat.format("call report_activeratio_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_ACTIVERATIO_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
						
			// Rookie Performance
			sqlcommand = MessageFormat.format(
					"call report_segmentation_rookie_performance_tiedagency_north(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

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

	public void fetchDataForSouthSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 56;
		final int SECTOR_ENDINGMP_ROWINDEX = 18;
		final int SECTOR_RECRUITMENT_ROWINDEX = 29;
		final int SECTOR_ALRECRUITMENTKPIs_ROWINDEX = 39;
		final int SECTOR_APE_ROWINDEX = 47;
		final int SECTOR_FYP_ROWINDEX = 195;
		final int SECTOR_RYP_ROWINDEX = 183;
		final int SECTOR_APE_CONTRIBUTION_ROWINDEX = 61;
		final int SECTOR_MANPOWER_ROWINDEX = 74;
		final int SECTOR_ACTIVE_ROWINDEX = 86;
		final int SECTOR_CASE_ROWINDEX = 110;
		final int SECTOR_CASESIZE_ROWINDEX = 122;
		final int SECTOR_CASEPERACTIVE_ROWINDEX = 134;
		final int SECTOR_APEPERMANPOWER_ROWINDEX = 146;
		final int SECTOR_APEPERACTIVE_ROWINDEX = 158;
		final int SECTOR_ACTIVERATIO_ROWINDEX = 98;
		final int SECTOR_ROOKIEPERFORMANCE_ROWINDEX = 170;
		final String SHEET_NAME = "South";

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
			CellStyle cellStylePercentage = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));
			cellStylePercentage.setDataFormat(createHelper.createDataFormat().getFormat("0.0%"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_segmentation_man_power_by_designation_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// Recruitment
			sqlcommand = MessageFormat.format(
					"call report_segmentation_recruitment_by_designation_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RECRUITMENT_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);

			// AL recruitment KPIs
			sqlcommand = MessageFormat.format(
					"call report_segmentation_al_recruitment_kpis_tiedagency_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ALRECRUITMENTKPIs_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

			// APE
			sqlcommand = MessageFormat.format("call report_ape_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APE_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// fyp
			sqlcommand = MessageFormat.format("call report_fyp_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_FYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// ryp
			sqlcommand = MessageFormat.format("call report_ryp_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_RYP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// APE contribution
			sqlcommand = MessageFormat.format("call report_ape_contribution_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_APE_CONTRIBUTION_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

			// manpower
			sqlcommand = MessageFormat.format("call report_manpower_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_MANPOWER_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// active
			sqlcommand = MessageFormat.format("call report_active_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_ACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// case
			sqlcommand = MessageFormat.format("call report_case_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASE_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// casesize
			sqlcommand = MessageFormat.format("call report_casesize_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASESIZE_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// case per active
			sqlcommand = MessageFormat.format("call report_caseperactive_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_CASEPERACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// APE per active
			sqlcommand = MessageFormat.format("call report_apeperactive_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERACTIVE_ROWINDEX, SECTOR_COLUMNINDEX, rs);			
			
			// APE per manpower
			sqlcommand = MessageFormat.format("call report_apepermanpower_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_APEPERMANPOWER_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);
					
			// active ratio
			sqlcommand = MessageFormat.format("call report_activeratio_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStylePercentage, SHEET_NAME, SECTOR_ACTIVERATIO_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);
			
			// Rookie Performance
			sqlcommand = MessageFormat.format(
					"call report_segmentation_rookie_performance_tiedagency_south(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle3, SHEET_NAME, SECTOR_ROOKIEPERFORMANCE_ROWINDEX,
					SECTOR_COLUMNINDEX, rs);

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

	public void fetchDataForEndingMPStructureSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Ending MP_Structure";

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
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_segmentation_man_power_by_designation_all_group(\"{0}\", \"{1}\");", inputPeriodFrom,
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

	public void fetchDataForRecruitmentStructureSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Recruitment_Structure";

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
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_segmentation_recruitment_by_designation_all_group(\"{0}\", \"{1}\");", inputPeriodFrom,
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

	public void fetchDataForRecruitmentKPIStructureSheet(String excelTemplate, String excelReport,
			String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Recruitment KPI_Structure";

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
			CellStyle cellStyle1 = book.createCellStyle();
			CellStyle cellStyle2 = book.createCellStyle();
			CellStyle cellStyle3 = book.createCellStyle();
			CreationHelper createHelper = book.getCreationHelper();
			cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// ending manpower
			sqlcommand = MessageFormat.format(
					"call report_segmentation_al_recruitment_kpis_all_group_level(\"{0}\", \"{1}\");", inputPeriodFrom,
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

	public void fetchDataForRookieMetricSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Rookie Metric";

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
					"call report_segmentation_rookie_performance_tiedagency_all_group(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
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

	public void fetchDataForGASheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "GA";

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

	public void fetchDataForRiderSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 3;
		final String SHEET_NAME = "Rider";

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
			sqlcommand = MessageFormat.format("call report_rider_attach_tiedagency_all_group(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
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

	public void fetchDataForProductMixSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = 0;
		final int SECTOR_APEMIX_ROWINDEX = 4;
		final int SECTOR_APERIDERMIX_ROWINDEX = 17;
		final int SECTOR_COUNTRIDERMIX_ROWINDEX = 34;
		final String SHEET_NAME = "Product Mix";

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

			// report_product_mix_tiedagency
			sqlcommand = MessageFormat.format("call report_product_mix_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APEMIX_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// report_product_mix_rider_tiedagency
			sqlcommand = MessageFormat.format("call report_product_mix_rider_tiedagency(\"{0}\", \"{1}\");",
					inputPeriodFrom, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_APERIDERMIX_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);

			// report_product_mix_rider_tiedagency
			sqlcommand = MessageFormat.format(
					"call report_product_mix_count_rider_products_tiedagency(\"{0}\", \"{1}\");", inputPeriodFrom,
					inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, SHEET_NAME, SECTOR_COUNTRIDERMIX_ROWINDEX, SECTOR_COLUMNINDEX,
					rs);

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
		} catch (SQLException e) {
			e.printStackTrace();
			logger.error(sqlcommand);
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

	public void fetchDataForBDSheet(String excelTemplate, String excelReport, String inputPeriodFrom,
			String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "BD";

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
			sqlcommand = MessageFormat.format("call report_team_performance(\"{0}\", \"{1}\");", inputPeriodFrom,
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

	public void fetchDataForRetentionSheet(String excelTemplate, String excelReport, String inputPeriodTo) {
		int inputColumIndex = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 8;
		final String SHEET_NAME = "5.0 AG retention";
		final int THE_DATE_COLUMN_INDEX = 1;
		FileInputStream fis = null;
		XSSFWorkbook book = null;
		FileOutputStream fos = null;
		MySQLConnect mySQLConnect = null;
		String sqlcommand = null;
		ResultSet rs = null;
		Iterator<Row> iterator = null;
		// format date
		SimpleDateFormat sdfMMMyy = new SimpleDateFormat("MMM-yy");

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			// CellStyle cellStyle1 = book.createCellStyle();
			 CellStyle cellStyle2 = book.createCellStyle();
			// CellStyle cellStyle3 = book.createCellStyle();
			 CreationHelper createHelper = book.getCreationHelper();
			 cellStyle2.setDataFormat(createHelper.createDataFormat().getFormat("#,##0"));
			// cellStyle3.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.0"));

			// fetch data from the database
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			mySQLConnect.connect(true);

			// fetching retention data
			sqlcommand = MessageFormat.format("call report_retention(\"{0}\", \"{1}\");", inputPeriodTo, inputPeriodTo);
			rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
			
			// reading sheet retention data
			XSSFSheet sheet = book.getSheet(SHEET_NAME);
			// while loop: finds the column index that its date value match with inputPeriodTo
			iterator = sheet.iterator();
			while_loop: while (iterator.hasNext()) {
				Row row = iterator.next();
				for (Cell cell: row) {
					// excel stores data value as a numeric
					if (cell.getCellTypeEnum() == CellType.NUMERIC) {
						// check if a numeric is a valid excel date
						if (DateUtil.isValidExcelDate(cell.getNumericCellValue())) {
							Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());
							if (sdfMMMyy.format(javaDate).equals(rs.getMetaData().getColumnName(2))) {
								inputColumIndex = cell.getColumnIndex();
								break while_loop;
							}
						}
					}
				}
			}

			if (inputColumIndex != -1) {
				iterator = sheet.iterator();
				while (iterator.hasNext()) {
					Row row = iterator.next();
					Cell cell = row.getCell(THE_DATE_COLUMN_INDEX);
					if (null != cell) {
						CellType cellTypeEnum = cell.getCellTypeEnum();
						switch (cellTypeEnum) {
						case NUMERIC:
							// check if the double value in the cell is a valid excel date
							if (DateUtil.isValidExcelDate(cell.getNumericCellValue())) {
								// convert excel date (in double value) to java date
								Date javaDate = DateUtil.getJavaDate(cell.getNumericCellValue());

								// with each date vale of the excel row, I check if the date value is same with
								// joining_month value from
								// the result set. If it is, I find the appropriate column to fill data
								if (!rs.isFirst())
									rs.first();
								while (rs.next()) {
									if (sdfMMMyy.format(javaDate)
											.equalsIgnoreCase(rs.getString(THE_DATE_COLUMN_INDEX))) {
										Cell newcell = row.getCell(inputColumIndex);
										if (null == newcell) {
											newcell = row.createCell(inputColumIndex);
										}
										newcell.setCellValue(rs.getDouble(2));
										newcell.setCellStyle(cellStyle2);
										break;
									}
								}

								break;
							}

						case STRING:
							System.out.println(cell.getStringCellValue());
							break;
						default:
							break;
						}
					}
				}
			}

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
		final int asAtRowIdx = 2;
		final int asAtColIdx = 6;
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
}
