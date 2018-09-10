package dbtool.msaccess2mysql;

import java.util.Arrays;
import java.util.Properties;

import org.apache.log4j.Logger;

import main.utils.Utils;

/**
 * Hello world!
 *
 */
public class bancassurance {
	final static Logger logger = Logger.getLogger(bancassurance.class);
	public static void main(String[] args) {
		Properties systemProperties = Utils.loadSystemProperties("config_msaccess_file.properties");
		String m0Start = systemProperties.getProperty("m0Start");
		
		MsAccess2MySQL msAccess2MySQL = new MsAccess2MySQL(systemProperties.getProperty("KPI_BANCA")); 
//		MsAccess2MySQL msAccess2MySQL = new MsAccess2MySQL("E:\\DA_DATA\\KPI_BANCA.accdb"); 
		
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("KPITotal", "raw_kpitotal_bancassurance_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
	
		msAccess2MySQL.runStoreProcedureOnMySQL("call job_update_data_bancassurance('" + m0Start + "');");
		msAccess2MySQL.closeAllConnections();
		
		if(logger.isInfoEnabled()){
			logger.info("=== DONE ===");
		}
	}
}
