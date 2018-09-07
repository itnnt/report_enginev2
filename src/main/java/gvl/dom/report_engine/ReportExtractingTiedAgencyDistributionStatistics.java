package gvl.dom.report_engine;

import java.util.Properties;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.TiedAgencyDistributionStatistics;
import gvl.dom.report_engine.reports.TiedAgencyPerformanceSegmentReport;
import main.utils.Utils;

public class ReportExtractingTiedAgencyDistributionStatistics {
	final static Logger logger = Logger.getLogger(ReportExtractingTiedAgencyDistributionStatistics.class);
	public static void main(String[] args) {
		Properties systemProperties = Utils.loadSystemProperties("config_distribution_statistics.properties");
		String excelTemplate = systemProperties.getProperty("excelTemplate");
		String excelReport = systemProperties.getProperty("excelReport");
		String excelReport2 = systemProperties.getProperty("excelReport_sa_excluded");
		String y0 = systemProperties.getProperty("y0");
		String m0End = systemProperties.getProperty("m0End");
		
//		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\bin\\Distribution Stat_GVL_template.xlsx";
//		String excelReport = "E:\\eclipse-workspace\\report_engine\\bin\\out\\Distribution Stat_GVL_2018-07-31.xlsx";
//		String excelReport2 = "E:\\eclipse-workspace\\report_engine\\bin\\out\\Distribution Stat_GVL_SA_excl_2018-07-31.xlsx";
//		String y0 = "2018-01-01";
//		String m0End = "2018-07-31";
		
		TiedAgencyDistributionStatistics dis = new TiedAgencyDistributionStatistics();
		dis.fetchDataForReport(excelTemplate, excelReport, excelReport2, y0, m0End);
		
		if(logger.isInfoEnabled()){
			logger.info("=== DONE ===");
		}
	}

}
