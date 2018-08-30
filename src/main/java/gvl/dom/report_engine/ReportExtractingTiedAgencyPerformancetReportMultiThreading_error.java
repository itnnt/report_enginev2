package gvl.dom.report_engine;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.MessageFormat;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyPerformanceReport;
import main.utils.MySQLConnect;

public class ReportExtractingTiedAgencyPerformancetReportMultiThreading_error {
	final static Logger logger = Logger.getLogger(ReportExtractingTiedAgencyPerformancetReportMultiThreading_error.class);
	final static TiedAgencyPerformanceReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceReport();
	
	private static void startAThreadFetchingDataIntoExcel(final String sqlcommand, final String excelTemplate, final String excelReport, final String sheetname, final int col, final int row) {
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		Runnable r1 = new Runnable() {
			public void run() {
				if (logger.isInfoEnabled()) {
					logger.info("Fetching data for sheet: " + sheetname);
				}
				try {
					MySQLConnect mySQLConnect = null;
					ResultSet rs = null;
					// fetch data from the database
					mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
					mySQLConnect.connect(true);
					rs = mySQLConnect.runStoreProcedureToGetReturn(sqlcommand);
					tiedAgencyPerformanceSegmentReport.writeDataForSheet(excelTemplate, excelReport, sheetname, col, row, rs);
					mySQLConnect.close();
				} catch (SQLException e) {
					e.printStackTrace();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		};
		Thread thread1 = new Thread(r1);
		thread1.start();
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
	}
	
	public static void main(String[] args) {
		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_template.xlsm";
		final String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_PERFORMANCE_REPORT_dynamic_2018-07-31-RESULT.xlsm";
		
		// Ngày đầu tiên của năm liền trước năm hiện tại
		String y1 = "2017-01-01";
		
		// Ngày cuối cùng của năm liền trước năm hiện tại
		String y1End = "2017-12-31";
		
		// Ngày đầu tiên của năm hiện tại
		String y0 = "2018-01-01";
		
		// Ngày cuối cùng của năm hiện tại
		String y0End = "2018-12-31";
		
		// Ngày đầu tiên của tháng chạy report
		final String m0Start = "2018-07-01";
		
		// Ngày cuối cùng của tháng chạy report
		final String m0End = "2018-07-31";
		
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		if(logger.isInfoEnabled()){
			logger.info("Updating data for Cover sheet");
		}
		tiedAgencyPerformanceSegmentReport.updateCoverSheet(excelTemplate, excelReport, m0End);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		String sqlcommand = MessageFormat.format("call tiedagency_ape_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_in_segment_ape", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_mp_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_in_segment_manpower", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_casecount_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_in_segment_cscnt", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_active_report_dynamic(\"{0}\", \"{1}\" , 7);",m0Start, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_in_segment_active", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_ape(\"{0}\", \"{1}\" , 7);", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_seg_chart_y0_ape", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_mp(\"{0}\", \"{1}\" , 7);", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_seg_chart_y0_mp", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_case(\"{0}\", \"{1}\" , 7);", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_seg_chart_y0_case", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_active(\"{0}\", \"{1}\" , 7);", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_detail_seg_chart_y0_act", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_mp_tt(\"{0}\", \"{1}\" , 7);", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_manpower_monthly_total", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_sheet_data_detail_seg_chart_y0_mp_tt(\"{0}\", \"{1}\" , 7);", y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_manpower_monthly_total_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_recruitment_tt(\"{0}\", \"{1}\" );", y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "newrecruit_monthly_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_recruitment_tt(\"{0}\", \"{1}\" );", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "newrecruit_monthly", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_tt(\"{0}\", \"{1}\", 3 );", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_ape_y0", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_tt(\"{0}\", \"{1}\", 3 );", y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "data_ape_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_activeratio(\"{0}\", \"{1}\", 3 );", y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_active_ratio_monthly_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_activeratio(\"{0}\", \"{1}\", 3 );", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_active_ratio_monthly", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_casesize(\"{0}\", \"{1}\", 3 );",y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_casesize_monthly_12ms_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_casesize(\"{0}\", \"{1}\", 3 );",y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "group_casesize_monthly", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_case_per_active(\"{0}\", \"{1}\", 3 );", y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "case_per_active_monthly_12ms_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_case_per_active(\"{0}\", \"{1}\", 3 );", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "case_per_active_monthly", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_per_active(\"{0}\", \"{1}\", 3 );",y1, y1End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "ape_per_active_monthly_12ms_y1", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		sqlcommand = MessageFormat.format("call tiedagency_report_dynamic_ape_per_active(\"{0}\", \"{1}\", 3 );", y0, m0End);
		startAThreadFetchingDataIntoExcel(sqlcommand, excelReport, excelReport, "ape_per_active_monthly", -1, 0);
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		
		/* --------------------------------------------------------------------------------- */
		/* --------------------------------------------------------------------------------- */
		
	}

}
