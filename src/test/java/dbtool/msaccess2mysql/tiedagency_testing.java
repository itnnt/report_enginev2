package dbtool.msaccess2mysql;

import java.util.Arrays;
import java.util.Properties;

import main.utils.Utils;

/**
 * Hello world!
 *
 */
public class tiedagency_testing {

	public static void main(String[] args) {
		MsAccess2MySQL msAccess2MySQL = new MsAccess2MySQL("E:\\DA_DATA\\KPI_PRODUCTION.accdb"); 
		MsAccess2MySQL msAccess2MySQL_persistency = new MsAccess2MySQL("E:\\DA_DATA\\PERSISTENCY_PRODUCTION.accdb"); 
		
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("GATotal", "raw_gatotal_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("Manpower_Active Ratio", "raw_manpower_active_ratio_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("KPITotal", "raw_kpitotal_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("Adlist", "raw_adlist_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("AgentList", "raw_agentlist_profile_update",	Arrays.asList(new String[] {}), Arrays.asList(new String[] {"channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name", "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("AgentList", "raw_agentlist_hierarchy_update",	Arrays.asList(new String[] {"mastercd", "agentcd", "channelcd", "channel_name", "subchannelcd", "subchannel_name", "agencycd", "agency_name", "point_of_sale_cd", "point_of_sale_name", "reference_no", "reportingcd", "reporting_name", "unitcd_of_reportingcd", "unit_name_of_reportingcd", "introducercd", "introducer_name", "recuitercd", "recuiter_name", "mentorcd", "mentor_name", "joining_date", "approved_date", "agent_designation", "staffcd", "agent_status", "classification", "termination_date", "termination_reason", "reinstatement_date", "reinstatement_reason" , "regioncd", "region_name", "regionheadcd", "region_head_name", "zonecd", "zone_name", "zone_head_cd", "zone_head_name", "teamcd", "team_name", "teamheadcd", "team_head_name", "officecd", "office_name", "officeheadcd", "officeheadname", "bmcd", "bm_name", "branchcd", "branch_name", "umcd", "um_name", "unitcd", "unit_name"}),	Arrays.asList(new String[] {}) 		);
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("Genlion_Report", "raw_genlion_report_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL.createOrAlterTableThenTruncateBeforeInsert("Target", "raw_target_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		
		msAccess2MySQL_persistency.createOrAlterTableThenTruncateBeforeInsert("persistency", "raw_persistency_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL_persistency.createOrAlterTableThenTruncateBeforeInsert("persistency_y1", "raw_persistency_y1_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		msAccess2MySQL_persistency.createOrAlterTableThenTruncateBeforeInsert("persistency_y2", "raw_persistency_y2_update", Arrays.asList(new String[] {}), Arrays.asList(new String[] {}));
		
		msAccess2MySQL.runStoreProcedureOnMySQL("call job_update_data_tiedagency('2018-08-01');");
		msAccess2MySQL.closeAllConnections();
	}
}
