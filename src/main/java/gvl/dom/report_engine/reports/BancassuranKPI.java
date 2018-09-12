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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
public class BancassuranKPI extends XSSFWorkbookReport {
	final static Logger logger = Logger.getLogger(BancassuranKPI.class);
	
	DataSource dataSource;
	ConnectionPool jdbcObj;
	public BancassuranKPI() {
		jdbcObj = new ConnectionPool();
		try {
			dataSource = jdbcObj.setUpPool();
			jdbcObj.printDbStatus();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void fetchDataKPI(String excelTemplate, String excelReport, String sheetname, String year) {
		final int SECTOR_COLUMNINDEX = -1;
		final int SECTOR_ENDINGMP_ROWINDEX = 0;

		FileInputStream fis = null;
		XSSFWorkbook book = null;
		SXSSFWorkbook sbook = null;
		FileOutputStream fos = null;
		
		String sqlcommand = null;
		ResultSet rs = null; Connection connObj = null;	PreparedStatement pstmt = null;

		try {
			// open the template
			File fileTemplate = new File(excelTemplate);
			fis = new FileInputStream(fileTemplate);
			book = new XSSFWorkbook(fis);
			sbook = new SXSSFWorkbook(book, 100);
//			sbook = new SXSSFWorkbook();
			CellStyle cellStyle2 = null;

			// fetch data from the database
			connObj = dataSource.getConnection();

			sqlcommand = MessageFormat.format("call get_banca_kpi(\"{0}\");", year);
			pstmt = connObj.prepareStatement(sqlcommand); 
			rs = pstmt.executeQuery();
			// write the result set to the excel
//			XLSXReadWriteHelper.write(book, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);
			XLSXReadWriteHelper.writeLargeFile(sbook, cellStyle2, sheetname, SECTOR_ENDINGMP_ROWINDEX, SECTOR_COLUMNINDEX, rs);

			// write to the new file
			File fileSavedTo = new File(excelReport);
			// open an OutputStream to save written data into Excel file
			fos = new FileOutputStream(fileSavedTo);
			sbook.write(fos);

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
				
				 // dispose of temporary files backing this workbook on disk
		        sbook.dispose();
		        sbook.close();
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
