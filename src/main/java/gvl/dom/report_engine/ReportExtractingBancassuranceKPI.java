package gvl.dom.report_engine;

import java.util.Properties;

import org.apache.log4j.Logger;

import gvl.dom.report_engine.reports.BancassuranKPI;
import main.utils.Utils;

public class ReportExtractingBancassuranceKPI {
	final static Logger logger = Logger.getLogger(ReportExtractingBancassuranceKPI.class);
	public static void main(String[] args) {
		Properties systemProperties = Utils.loadSystemProperties("config_bancassurance_kpi.properties");
		String excelTemplate = systemProperties.getProperty("excelTemplate");
		String excelReport = systemProperties.getProperty("excelReport");
		String y0 = systemProperties.getProperty("y0");
		
//		String excelTemplate = "E:\\eclipse-workspace\\report_engine\\bin\\bancassurance_kpis_template.xlsx";
//		String excelReport = "E:\\eclipse-workspace\\report_engine\\bin\\out\\bancassurance_kpis.xlsx";
//		String y0 = "2018";
		
		BancassuranKPI dis = new BancassuranKPI();
		dis.fetchDataKPI(excelTemplate, excelReport, "KPI", y0);
		
		if(logger.isInfoEnabled()){
			logger.info("=== DONE ===");
		}
	}

}
