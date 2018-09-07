package gvl.dom.report_engine.reports;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;

import javax.sql.DataSource;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import main.utils.ConnectionPool;
import main.utils.MySQLConnect;
import main.utils.XLSXReadWriteHelper;

/**
 * Hello world!
 *
 */
public class TiedAgencyDistributionStatistics extends XSSFWorkbookReport {
	final static Logger logger = Logger.getLogger(TiedAgencyDistributionStatistics.class);
	
	DataSource dataSource;
	ConnectionPool jdbcObj;
	public TiedAgencyDistributionStatistics() {
		jdbcObj = new ConnectionPool();
		try {
			dataSource = jdbcObj.setUpPool();
			jdbcObj.printDbStatus();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void fetchDataForReport(String excelTemplate, String excelReport, String excelReport2, String inputPeriodFrom, String inputPeriodTo) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		FileInputStream fis2 = null;
		XSSFWorkbook book = null;
		XSSFWorkbook book2 = null;
		FileOutputStream fos = null;
		FileOutputStream fos2 = null;
		
		String sqlcommand = null;
		ResultSet rs = null; Connection connObj = null;	PreparedStatement pstmt = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			fis2 = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			book2 = new XSSFWorkbook(fis2);
			CellStyle cellStyle2 = null;

			// fetch data from the database
			connObj = dataSource.getConnection();
			// ---------------------------------------------------------------------------------
			String sheetName = "data_new_case";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_casecount(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			rs.beforeFirst();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_newcase_newagent";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_casecount_newagent(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			rs.beforeFirst();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_agent_and_leader";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power_sa_excl(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_agent";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power_ag_plus_us(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power_ag_plus_us_sa_excl(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_leader";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power_leader(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			sqlcommand = MessageFormat.format("call report_distribution_statistics_man_power_leader_sa_excl(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_newagent_recruit";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_newagent_recruit(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			rs.beforeFirst();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_active_leader";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_active_leader(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			rs.beforeFirst();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_active_agent";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call report_distribution_statistics(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			sqlcommand = MessageFormat.format("call report_distribution_statistics_sa_excl(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_retention_rate";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_retentionrate_v1(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write the result set to the excel
			rs.beforeFirst();
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_retention_rate_branch";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_retentionrate_bra(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write the result set to the excel
			rs.beforeFirst();
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// ---------------------------------------------------------------------------------
			sheetName = "data_retention_rate_ga";
			if(logger.isInfoEnabled()){
				logger.info("Fetching data for: " + sheetName);
			}
			sqlcommand = MessageFormat.format("call tiedagency_report_distributionstatictics_retentionrate_ga(\"{0}\", \"{1}\");", inputPeriodFrom, inputPeriodTo);
			pstmt = connObj.prepareStatement(sqlcommand); rs = pstmt.executeQuery();
			// write the result set to the excel
			XLSXReadWriteHelper.write(book, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			
			// write the result set to the excel
			rs.beforeFirst();
			XLSXReadWriteHelper.write(book2, cellStyle2, sheetName, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			// ---------------------------------------------------------------------------------
			
			// write to the new file
			File fileSavedTo = new File(excelReport);
			File fileSavedTo2 = new File(excelReport2);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			fos2 = new FileOutputStream(fileSavedTo2);
			book.write(fos);
			book2.write(fos2);

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
				fos2.close();
				book.close();
				fis.close();
				fis2.close();
				if(rs != null) {
					rs.close();
				}
				// Closing PreparedStatement Object
				if(pstmt != null) {
					pstmt.close();
				}
				// Closing Connection Object
				if(connObj != null) {
					connObj.close();
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
	}

}
