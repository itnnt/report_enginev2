package dbtool.msaccess2mysql;

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Set;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;

import com.healthmarketscience.jackcess.Column;
import com.healthmarketscience.jackcess.DataType;
import com.healthmarketscience.jackcess.Row;
import com.healthmarketscience.jackcess.Table;

import dbtool.msaccess2mysql.util.MsAccessConnect;
import main.utils.CommonUtil;
import main.utils.MySQLConnect;

public class MsAccess2MySQL {
	final static Logger logger = Logger.getLogger(MsAccess2MySQL.class);
	public enum TypeMapping {
		/* Ms Access type mapped to MySQL type */
		LONG("int"),
		DOUBLE("double"),
		TEXT("varchar(255)"),
		SHORT_DATE_TIME("DATE"),
	    UNKNOWN("");
	    private String mySqlType;
	    TypeMapping(String mySqlType) {
	        this.mySqlType = mySqlType;
	    }
	    public String mySqlType() {
	        return mySqlType;
	    }
	}
	
	CommonUtil util = new CommonUtil();
	MsAccessConnect msAccessConnect;
	MySQLConnect mySQLConnect;
	
	public MsAccess2MySQL(String msaccessDbFile) {
		super();
		try {
			msAccessConnect = new MsAccessConnect(msaccessDbFile);
			mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
			/* connect to source db and destination db */
			msAccessConnect.connect();
			mySQLConnect.connect(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void synchData() {
		try {
			
			Set<String> tableNames = msAccessConnect.getTableNames();
			for (String t: tableNames) {
				System.out.println("Processing table " + t + "-----------------");
				if (util.removeSpecialChar(t).equals("agentlist")) {
					createOrAlterTableThenTruncateBeforeInsert(t, "raw_" + util.removeSpecialChar(t) + "_profile",	Arrays.asList(new String[] {}), Arrays.asList(new String[] {"channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name", "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}));
					checkSynchedTable(t, "raw_" + util.removeSpecialChar(t) + "_profile",	Arrays.asList(new String[] {}), Arrays.asList(new String[] {"channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name",  "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}));
				
					createOrAlterTableThenTruncateBeforeInsert(t, "raw_" + util.removeSpecialChar(t) + "_hierarchy",	Arrays.asList(new String[] {"mastercd", "agentcd", "channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name", "joining_date", "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}),	Arrays.asList(new String[] {}) 		);
					checkSynchedTable(t, "raw_" + util.removeSpecialChar(t) + "_hierarchy",	Arrays.asList(new String[] {"mastercd", "agentcd", "channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name", "joining_date", "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}),	Arrays.asList(new String[] {}) 		);
				} else {
					createOrAlterTableThenTruncateBeforeInsert(t, "raw_" + util.removeSpecialChar(t) + "", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
					checkSynchedTable(t, "raw_" + util.removeSpecialChar(t) + "", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	/**
	 * Check destination table: if it not exist, create it. 
	 * If the table is exist, compare source table's colums and destination table's column: alter new column for the destination table if any
	 * TRUNCATE add the data from the destination table
	 * Copy data from the source table to the destination table
	 */
	public void createOrAlterTableThenTruncateBeforeInsert(String accessTablename, String mySqlTablename, List<String> includedFields, List<String> excludedFields) {
		try {
			/* get data from the access table */
			Table tbl = msAccessConnect.getTable(accessTablename);
			/* the 4 list below have to be the same position of each column */
			List<String> colNameColTypePairs =  new ArrayList<String>();
			List<String> colNames =  new ArrayList<String>();
			List<String> simpleColNames =  new ArrayList<String>();
			List<String> originalColNames =  new ArrayList<String>();
			/* get columns' names and columns' types */
			for (Column column : tbl.getColumns()) {
				String columnName = util.removeSpecialChar(column.getName());
				boolean theFieldIsIncluded = false;
				if (includedFields.size() == 0 && excludedFields.size() == 0) {
					theFieldIsIncluded = true;
				} else if (includedFields.size() > 0 && excludedFields.size() == 0) {
					theFieldIsIncluded = includedFields.contains(columnName);
				} else if (includedFields.size() == 0 && excludedFields.size() > 0) {
					theFieldIsIncluded = !excludedFields.contains(columnName);
				} else if (includedFields.size() > 0 && excludedFields.size() > 0) {
					theFieldIsIncluded = (includedFields.contains(columnName) && !excludedFields.contains(columnName));
				}
				
				if (theFieldIsIncluded) {
//					 System.out.println("`" + columnName + "`" + "\t" + column.getType() );
					colNameColTypePairs.add("`" + columnName + "`" + "\t" + TypeMapping.valueOf(column.getType().toString()).mySqlType);
					colNames.add("`" + columnName + "`");
					simpleColNames.add(columnName);
					originalColNames.add(column.getName());
				}
			}
			
			/* check if the table is exist */
			if (mySQLConnect.isTableExist(mySqlTablename)) {
				/* get columns' names from the exist table */
				ResultSet rs = mySQLConnect.getColumns(mySqlTablename);
				List<String> colNamesFromTheExistTable = new ArrayList<String>();
				while (rs.next()) {
					colNamesFromTheExistTable.add(rs.getString("COLUMN_NAME"));
				}
				/* find if there are any columns added */
				List<String> addedColumns = new ArrayList<String>(simpleColNames);
				addedColumns.removeAll( colNamesFromTheExistTable );
				for (String s : addedColumns) {
					System.out.println("Alter the table " + accessTablename + ": add columns " + s);
					/* alter the table so that it can be added the new columns */
					for (int i=0; i< simpleColNames.size(); i++) {
						/* find the position of the new columns in the list */
						if (simpleColNames.get(i).equals(s)) {
							mySQLConnect.addColumn(mySqlTablename, colNameColTypePairs.get(i));
						}
					}
				}
			} else {
				/* create new table in mysql db if the table is not exist */
				mySQLConnect.createTable(mySqlTablename, StringUtils.join(colNameColTypePairs, ","));
			}
			/* truncate data before inserting */
			mySQLConnect.truncate(mySqlTablename);
			/* get values and values' type, cook it: (v1,v2,v3),(v4,v5,v6)... */
			List<String> values = new ArrayList<String>();
			for (Row row : tbl) {
				List<String> tem = new ArrayList<String>();
				for (String columnName : originalColNames) {
					Column column = tbl.getColumn(columnName);
					Object value = row.get(columnName);
					if (value == null) {
						tem.add("null");
					} else if (value instanceof java.lang.String) {
						if (((java.lang.String) value).contains("'")) {
							value = ((java.lang.String) value).replace("'", "''");
						}
						if (((java.lang.String) value).contains("\\")) {
							value = ((java.lang.String) value).replace("\\", "\\\\");
						}
						tem.add("'" + value + "'");
					} else if (column.getType() == DataType.SHORT_DATE_TIME) {
						/* remove ICT from the string before converting */
						String temDate = value.toString();
						temDate = temDate.replace("ICT ", "");
						SimpleDateFormat dateFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss yyyy");
						Date temDateVal = dateFormat.parse(temDate);
						SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
						tem.add("'" + df.format(temDateVal) + "'");
					} else {
						tem.add(value.toString());
					}
				}
				String join = StringUtils.join(tem, ",");
				values.add("(" + join + ")");
			}
			
			/* insert data */
			List<String> batchValues = new ArrayList<String>();
			int rowIndex  = 0;
			for (String v: values) {
				batchValues.add(v);
				rowIndex++;
				/* insert each time 1000 rows */
				if (rowIndex % 1000 == 0 || rowIndex == values.size()) {
					mySQLConnect.insert(mySqlTablename, StringUtils.join(colNames, ","), StringUtils.join(batchValues, ","));
					batchValues = new ArrayList<String>();
				}	
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
	}
	
	public void checkSynchedTable(String accessTablename, String mySqlTablename, List<String> includedFields, List<String> excludedFields) {
		logger.info(new String(new char[200]).replace("\0", "-"));
		try {
			/* get data from the access table */
			Table tbl = msAccessConnect.getTable(accessTablename);
			/* the 4 list below have to be the same position of each column */
			List<String> colNameColTypePairs =  new ArrayList<String>();
			List<String> colNames =  new ArrayList<String>();
			List<String> simpleColNames =  new ArrayList<String>();
			List<String> originalColNames =  new ArrayList<String>();
			/* get columns' names and columns' types */
			for (Column column : tbl.getColumns()) {
				String columnName = util.removeSpecialChar(column.getName());
				boolean theFieldIsIncluded = false;
				if (includedFields.size() == 0 && excludedFields.size() == 0) {
					theFieldIsIncluded = true;
				} else if (includedFields.size() > 0 && excludedFields.size() == 0) {
					theFieldIsIncluded = includedFields.contains(columnName);
				} else if (includedFields.size() == 0 && excludedFields.size() > 0) {
					theFieldIsIncluded = !excludedFields.contains(columnName);
				} else if (includedFields.size() > 0 && excludedFields.size() > 0) {
					theFieldIsIncluded = (includedFields.contains(columnName) && !excludedFields.contains(columnName));
				}
				
				if (theFieldIsIncluded) {
					colNameColTypePairs.add("`" + columnName + "`" + "\t" + TypeMapping.valueOf(column.getType().toString()).mySqlType);
					colNames.add("`" + columnName + "`");
					simpleColNames.add(columnName);
					originalColNames.add(column.getName());
				}
			}
			
			/* check if the table is exist */
			if (mySQLConnect.isTableExist(mySqlTablename)) {
				/* get columns' names from the exist table */
				ResultSet rs = mySQLConnect.getColumns(mySqlTablename);
				List<String> colNamesFromTheExistTable = new ArrayList<String>();
				while (rs.next()) {
					colNamesFromTheExistTable.add(rs.getString("COLUMN_NAME"));
				}
				/* find if there are any columns added */
				List<String> addedColumns = new ArrayList<String>(simpleColNames);
				addedColumns.removeAll( colNamesFromTheExistTable );
				if (addedColumns.size() == 0) {
					logger.info("Table " + accessTablename + " - " + mySqlTablename + " have the same number of columns");
				} else {
					logger.error("Table " + accessTablename + " has more fields than " + mySqlTablename + "'s field. They are " + StringUtils.join(addedColumns, ","));
				}
			} else {
				logger.error("Table " + accessTablename + " has not have a copied table on MySQL database. It MUST HAVE a table " + mySqlTablename + "created on MySQL DB");
			}
			
			/* compare number of rows on the two table */
			int rc1 = tbl.getRowCount();
			int rc2 = mySQLConnect.getRowCount(mySqlTablename);
			if (rc1 == rc2) {
				logger.info("Table " + accessTablename + " - " + mySqlTablename + " have the same number of rows: " + rc1);
			} else {
				logger.error("Table " + accessTablename + " has " + rc1 + " rows " + " - " + mySqlTablename + " has " + rc2 + " rows.");
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
		logger.info(new String(new char[200]).replace("\0", "-"));
	}
	
	public void runStoreProcedureOnMySQL(String callStoreProcedureCommand) {
		try {
			mySQLConnect.runStoreProcedure(callStoreProcedureCommand);
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
	
	public void closeAllConnections () {
		try {
			msAccessConnect.close();
			mySQLConnect.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void mappingTables() {
		MsAccessConnect msAccessConnect = new MsAccessConnect("D:\\Data\\DA_201712\\KPI_PRODUCTION_20171231.accdb");
		MySQLConnect mySQLConnect = new MySQLConnect("localhost", 3306, "root", "root", "generali");
		try {
			msAccessConnect.connect();
			Set<String> tables = msAccessConnect.getTableNames();
//			for (String t: tables) {
//				System.out.println(t);
//			}
			
			
			Table tbl = msAccessConnect.getTable("KPITotal");
			System.out.println(tbl.getRowCount());
//			for (Row row : tbl) {
//				Object[] array = row.values().toArray();
//				for (Object o: array) {
//					System.out.println(o);				
//				}
//			}
			
			
			for (Row row : tbl) {
				for (Column column : tbl.getColumns()) {
					String columnName = column.getName();
					Object value = row.get(columnName);
//					System.out.println("Column " + columnName + "(" + column.getType() + "): " + value + " ("
//							+ value.getClass() + ")");
				}
			}
			
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} finally {
			try {
				System.out.println("Closed db");
				msAccessConnect.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

//		// Get all tables names and add them to combobox
//		try {
//			Set<String> tables = db.getTableNames();
//
//			for (String t : tables) {
//				System.out.println(t);
//				Table tbl = db.getTable(t);
//				for (Row row : tbl) {
//					System.out.println(row.values().toArray());
//				}
//			}
//		} catch (Exception e) {
//		}
		
		
		
		try {
			mySQLConnect.connect(false);
			mySQLConnect.select();
			System.out.println(mySQLConnect.isTableExist("product"));
			System.out.println(mySQLConnect.isTableExist("prsoduct"));
			mySQLConnect.createTable("exampleTable2");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
}
