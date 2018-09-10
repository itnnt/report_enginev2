package gvl.dom.report_engine;

import java.util.Properties;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyPerformanceSegmentReport;
import main.utils.Utils;

public class ReportExtractingTiedAgencyPerformanceSegmentReport {
	final static Logger logger = Logger.getLogger(ReportExtractingTiedAgencyPerformanceSegmentReport.class);
	public static void main(String[] args) {
		Properties systemProperties = Utils.loadSystemProperties("config_report_tiedagency_kpi_segmentation.properties");
		String excelTemplate = systemProperties.getProperty("excelTemplate");
		String excelReport = systemProperties.getProperty("excelReport");
		String y1 = systemProperties.getProperty("y1");
		String y0 = systemProperties.getProperty("y0");
		String y0End = systemProperties.getProperty("y0End");
		String m0End = systemProperties.getProperty("m0End");
		Integer sectorColumnIndex = Integer.valueOf(systemProperties.getProperty("SECTOR_COLUMNINDEX"));
		
//		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_template.xlsx";
//		String excelReport = "E:\\eclipse-workspace\\report_engine\\src\\main\\resources\\MONTHLY_AGENCY_SEGMENTATION_REPORT_2018-07-31-RESULT.xlsx";
//		String y1 = "2017-01-01";
//		String y0 = "2018-01-01";
//		String y0End = "2018-12-31";
//		String m0End = "2018-07-31";
		
		TiedAgencyPerformanceSegmentReport tiedAgencyPerformanceSegmentReport = new TiedAgencyPerformanceSegmentReport();
		tiedAgencyPerformanceSegmentReport.SECTOR_COLUMNINDEX = sectorColumnIndex;
		if(logger.isInfoEnabled()){
			logger.info("Updating data for Cover sheet");
		}
		tiedAgencyPerformanceSegmentReport.updateCoverSheet(excelTemplate, excelReport, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Country sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForCountrySheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for North sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForNorthSheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for South sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForSouthSheet(excelReport, excelReport, y0, m0End);

		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Ending MP_Structure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForEndingMPStructureSheet(excelReport, excelReport, y1, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RecruitmentStructure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRecruitmentStructureSheet(excelReport, excelReport, y1, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RecruitmentKPIStructure sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRecruitmentKPIStructureSheet(excelReport, excelReport, y1, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for RookieMetric sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRookieMetricSheet(excelReport, excelReport, y1, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for GA sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForGASheet(excelReport, excelReport, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for Rider sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRiderSheet(excelReport, excelReport, y1, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for ProductMix sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForProductMixSheet(excelReport, excelReport, y0, y0End);
		
		if(logger.isInfoEnabled()){
			logger.info("Fetching data for BD sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForBDSheet(excelReport, excelReport, y0, m0End);
		
		if (logger.isInfoEnabled()) {
			logger.info("Fetching data for Agent_retention sheet");
		}
		tiedAgencyPerformanceSegmentReport.fetchDataForRetentionSheet(excelReport, excelReport, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("=== DONE ===");
		}
	}

}
