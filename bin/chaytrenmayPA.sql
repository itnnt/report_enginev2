CREATE DEFINER=`root`@`localhost` PROCEDURE `job_update_data_tiedagency`(in param_business_date date)
BEGIN

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_genlion_report where str_to_date(`bussinessdate`, '%Y%m%d') >= param_business_date; 
insert into raw_genlion_report(
zone_name, office_name, team_name, team_head_name, bm_name, um_name, agentcd, agent_full_name, sumoffyp, sumofcase, per_y1, per_y2, xephang, finaldate, appeal, created_date, updated_date, bussinessdate 
)
select 
zone_name, office_name, team_name, team_head_name, bm_name, um_name, agentcd, agent_full_name, sumoffyp, sumofcase, per_y1, per_y2, xephang, finaldate, appeal, created_date, updated_date, date_format(last_day(param_business_date), '%Y%m%d') as bussinessdate
from raw_genlion_report_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency where `date`>=param_business_date;
insert into raw_persistency(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
FROM raw_persistency_update 
where `date`>=param_business_date
;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency_y1 where `date`>=param_business_date;
insert into raw_persistency_y1(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
FROM raw_persistency_y1_update 
where `date`>=param_business_date;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency_y2 where `date`>=param_business_date;
insert into raw_persistency_y2(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date` 
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date` 
FROM raw_persistency_y2_update 
where `date`>=param_business_date;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
truncate table raw_adlist;
insert into raw_adlist (
team_code, team_name, team_head_code, team_head_name, office_code, office_name, office_head_code, office_head_name, zone_code, zone_name, zone_head_code, zone_head_name, region_code, region_name, region_head_code, region_head_name, agency_name, agency_code, sub_channel_name, sub_channel_code, channel_name, channel_code, active, created_date, updated_date
)
select 
team_code, team_name, team_head_code, team_head_name, office_code, office_name, office_head_code, office_head_name, zone_code, zone_name, zone_head_code, zone_head_name, region_code, region_name, region_head_code, region_head_name, agency_name, agency_code, sub_channel_name, sub_channel_code, channel_name, channel_code, active, created_date, updated_date 
from raw_adlist_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_agentlist_profile
truncate table raw_agentlist_profile;
insert into raw_agentlist_profile (
mastercd, agentcd, agent_full_name, agent_first_name, applicationcd, application_date, joining_date, gender, title, date_of_birth, married_, education_level, course, major, institution, occupation, id_number, id_date, id_place, mobile_phone, home_phone, home_extension, business_phone, business_extension, email, res_address_province, res_address_district, res_address_ward, res_address_house_no_street, con_address_province, con_address_district, con_address_ward, con_address_house_no_street, tem_address_province, tem_address_district, tem_address_ward, tem_address_house_no_street, bank_name, bank_branch_name, bank_account, account_status, effective_date_from, pay_to, payment_mode, account, overridden_minimum_amount, taxid, mof_licence_no, mof_licence_eff_date_from, partfulltimecd, partfulltime_name, created_date, updated_date
) select
mastercd, agentcd, agent_full_name, agent_first_name, applicationcd, application_date, joining_date, gender, title, date_of_birth, married_, education_level, course, major, institution, occupation, id_number, id_date, id_place, mobile_phone, home_phone, home_extension, business_phone, business_extension, email, res_address_province, res_address_district, res_address_ward, res_address_house_no_street, con_address_province, con_address_district, con_address_ward, con_address_house_no_street, tem_address_province, tem_address_district, tem_address_ward, tem_address_house_no_street, bank_name, bank_branch_name, bank_account, account_status, effective_date_from, pay_to, payment_mode, account, overridden_minimum_amount, taxid, mof_licence_no, mof_licence_eff_date_from, partfulltimecd, partfulltime_name, created_date, updated_date
from raw_agentlist_profile_update
;

delete from raw_agentlist_hierarchy where business_date >= param_business_date; 
insert into raw_agentlist_hierarchy(
mastercd, agentcd, channelcd, channel_name, subchannelcd, subchannel_name, agencycd, agency_name, regioncd, region_name, regionheadcd, region_head_name, zonecd, zone_name, zone_head_cd, zone_head_name, teamcd, team_name, teamheadcd, team_head_name, officecd, office_name, officeheadcd, officeheadname, point_of_sale_cd, point_of_sale_name, bmcd, bm_name, branchcd, branch_name, reference_no, umcd, um_name, unitcd, unit_name, reportingcd, reporting_name, unitcd_of_reportingcd, unit_name_of_reportingcd, introducercd, introducer_name, recuitercd, recuiter_name, mentorcd, mentor_name, joining_date, approved_date, agent_designation, staffcd, agent_status, classification, termination_date, termination_reason, reinstatement_date, reinstatement_reason, created_date, updated_date, business_date
) select 
mastercd, agentcd, channelcd, channel_name, subchannelcd, subchannel_name, agencycd, agency_name, regioncd, region_name, regionheadcd, region_head_name, zonecd, zone_name, zone_head_cd, zone_head_name, teamcd, team_name, teamheadcd, team_head_name, officecd, office_name, officeheadcd, officeheadname, point_of_sale_cd, point_of_sale_name, bmcd, bm_name, branchcd, branch_name, reference_no, umcd, um_name, unitcd, unit_name, reportingcd, reporting_name, unitcd_of_reportingcd, unit_name_of_reportingcd, introducercd, introducer_name, recuitercd, recuiter_name, mentorcd, mentor_name, joining_date, approved_date, agent_designation, staffcd, agent_status, classification, termination_date, termination_reason, reinstatement_date, reinstatement_reason, created_date, updated_date, last_day(param_business_date) as business_date
from raw_agentlist_hierarchy_update;


-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_manpower_active_ratio
delete from raw_manpower_active_ratio where bussinessdate >= param_business_date;
insert into raw_manpower_active_ratio (
agent_code, agent_name, agent_designation, staffcd, classification, monthstart, monthend, active, activeex, activesp, activesp6m, ip, supervisor_code, supervisor_code_designation, unitcd, unit_name, branchcd, branch_name, teamcd, team_name, officecd, office_name, zonecd, zone_name, regioncd, region_name, channelcd, channel_name, joining_date, termination_date, reinstatement_date, bussinessdate, activefirstmonth, m1, created_date, updated_date
) select 
agent_code, agent_name, agent_designation, staffcd, classification, monthstart, monthend, active, activeex, activesp, activesp6m, ip, supervisor_code, supervisor_code_designation, unitcd, unit_name, branchcd, branch_name, teamcd, team_name, officecd, office_name, zonecd, zone_name, regioncd, region_name, channelcd, channel_name, joining_date, termination_date, reinstatement_date, bussinessdate, activefirstmonth, m1, created_date, updated_date
from raw_manpower_active_ratio_update
where bussinessdate >= param_business_date;

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_kpitotal
delete from raw_kpitotal where bussinessdate >= date_format(param_business_date, '%Y%m%d');
insert into raw_kpitotal (
bussinessdate, duedate, rdocnum, productcode, premium, billingfrequency, 
`share`, txncode, policyyear, premiumterm, agcode, agcodedesignation, agent_status, 
classification, supervisor_code, supervisor_codedesignation, unitcd, unit, 
branchcd, branch, teamcd, team, officecd, office, zonecd, zone, regioncd, region, 
channelcd, channel, subchannelcd, subchannel_name, sacstyp, ip, ape, `case`, 
fyp, 2yp, 3yp, 4yp, 5yp, ryp, ultp, ic, fyc, 2yc, 3yc, 4yc, 5yc, ryc, comrate, 
com, ipflcan, icflcan, fypflcan, fycflcan, ccflcan, ulbnflcan, cibn, ulbn, ulbn1, 
sumins, tranno, life, hissdte, hoissdte, hprrcvdt, bankkey, countriders, 
countridersreal, cibnn, cibnn1, cibnnflcan, agent_idnumber, policyowner_idnumber, 
crosssales, rider
)
select 
bussinessdate, duedate, rdocnum, productcode, premium, billingfrequency, 
`share`, txncode, policyyear, premiumterm, agcode, agcodedesignation, agent_status, 
classification, supervisor_code, supervisor_codedesignation, unitcd, unit, 
branchcd, branch, teamcd, team, officecd, office, zonecd, zone, regioncd, region, 
channelcd, channel, subchannelcd, subchannel_name, sacstyp, ip, ape, `case`, 
fyp, 2yp, 3yp, 4yp, 5yp, ryp, ultp, ic, fyc, 2yc, 3yc, 4yc, 5yc, ryc, comrate, 
com, ipflcan, icflcan, fypflcan, fycflcan, ccflcan, ulbnflcan, cibn, ulbn, ulbn1, 
sumins, tranno, life, hissdte, hoissdte, hprrcvdt, bankkey, countriders, 
countridersreal, cibnn, cibnn1, cibnnflcan, agent_idnumber, policyowner_idnumber, 
crosssales, rider
from raw_kpitotal_update where bussinessdate >= date_format(param_business_date, '%Y%m%d');

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
truncate table raw_gatotal;
insert into raw_gatotal (
region_name, officecd, office_name, gacode, ganame, agcode, agname, gatype, contractdt, effrom, effto, classification, housing, bonus12stend, q01, q02, q03, size, cashadv, targetlevel, openingdt, created_date, updated_date
)
select 
region_name, officecd, office_name, gacode, ganame, agcode, agname, gatype, contractdt, effrom, effto, classification, housing, bonus12stend, q01, q02, q03, size, cashadv, targetlevel, openingdt, created_date, updated_date
from raw_gatotal_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

delete from raw_target where teamcode not in ('NORTH', 'SOUTH', 'COUNTRY') ;
insert into raw_target (
teamname, teamcode, ape, activeag, newrecruit, newleader, newagent, fyp, txndate, created_date, updated_date
)
select
teamname, teamcode, ape, activeag, newrecruit, newleader, newagent, fyp, txndate, created_date, updated_date
from raw_target_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ------last run----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
call job_segment(param_business_date, last_day(param_business_date));
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

END




















alter table raw_agentlist_hierarchy add column segment varchar(50) null;
-- ---------------------------------------------------------------------------


CREATE DEFINER=`root`@`localhost` PROCEDURE `job_segment`(In fromdt date, in todt date)
BEGIN
call get_genlion(fromdt, todt,'tempgenlion');
call get_sa(fromdt, todt,'tempsa');

Alter table tempgenlion add index tempgenlion_index(business_date, agentcd);
ALTER TABLE tempsa ADD INDEX tempsa_index (business_date, agentcd );

update raw_agentlist_hierarchy a
left join tempgenlion b on a.agentcd=b.agentcd and last_day(a.business_date)=b.business_date
left join tempsa c on a.agentcd=c.agentcd and last_day(a.business_date)=c.business_date
set segment=case 
	when is_sa_agent is not null then 'SA'  
	when is_genlion_agent is not null then 'GEN LION (from Apr ''17)'  
	WHEN (last_day(a.joining_date) = last_day(a.business_date)) THEN 'Rookie in month'
	WHEN (last_day(a.joining_date + INTERVAL 1 MONTH) = last_day(a.business_date)) THEN 'Rookie last month'
	WHEN (last_day(a.business_date) BETWEEN last_day(a.joining_date + INTERVAL 2 MONTH) AND last_day(a.joining_date + INTERVAL 3 MONTH)) THEN '2-3 months'
	WHEN (last_day(a.business_date) BETWEEN last_day(a.joining_date + INTERVAL 4 MONTH) AND last_day(a.joining_date + INTERVAL 6 MONTH)) THEN '4-6 mths'
	WHEN (last_day(a.business_date) BETWEEN last_day(a.joining_date + INTERVAL 7 MONTH) AND last_day(a.joining_date + INTERVAL 12 MONTH)) THEN '7-12mth'
	WHEN (last_day(a.business_date) between last_day(a.joining_date + INTERVAL 13 MONTH) and now()) THEN '13+mth'
end 
where a.`business_date` between fromdt and todt
AND lower(a.`channelcd`) = 'tiedagency'
;


END

-- ---------------------------------------------------------------------------
call job_segment('2018-01-01', '2018-07-31');






CREATE DEFINER=`root`@`localhost` PROCEDURE `job_update_data_tiedagency`(in param_business_date date)
BEGIN

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_genlion_report where str_to_date(`bussinessdate`, '%Y%m%d') >= param_business_date; 
insert into raw_genlion_report(
zone_name, office_name, team_name, team_head_name, bm_name, um_name, agentcd, agent_full_name, sumoffyp, sumofcase, per_y1, per_y2, xephang, finaldate, appeal, created_date, updated_date, bussinessdate 
)
select 
zone_name, office_name, team_name, team_head_name, bm_name, um_name, agentcd, agent_full_name, sumoffyp, sumofcase, per_y1, per_y2, xephang, finaldate, appeal, created_date, updated_date, date_format(last_day(param_business_date), '%Y%m%d') as bussinessdate
from raw_genlion_report_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency where `date`>=param_business_date;
insert into raw_persistency(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
FROM raw_persistency_update 
where `date`>=param_business_date
;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency_y1 where `date`>=param_business_date;
insert into raw_persistency_y1(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date`
FROM raw_persistency_y1_update 
where `date`>=param_business_date;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
delete from raw_persistency_y2 where `date`>=param_business_date;
insert into raw_persistency_y2(
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date` 
)
select
agent_name, agent_code, totalapecurrent_indi, totalapeoriginal_indi, individual, totalapecurrent_unit, totalapeoriginal_unit, unit, totalapecurrent_bran, totalapeoriginal_bran, branch, totalapecurrent_office, totalapeoriginal_office, office, totalapecurrent_team, totalapeoriginal_team, team, totalapecurrent_zone, totalapeoriginal_zone, zone, totalapecurrent_region, totalapeoriginal_region, region, totalapecurrent_subchannel, totalapeoriginal_subchannel, subchannel, totalapecurrent_agency, totalapeoriginal_agency, agency, totalapecurrent_channel, totalapeoriginal_channel, `channel`, `date` 
FROM raw_persistency_y2_update 
where `date`>=param_business_date;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
truncate table raw_adlist;
insert into raw_adlist (
team_code, team_name, team_head_code, team_head_name, office_code, office_name, office_head_code, office_head_name, zone_code, zone_name, zone_head_code, zone_head_name, region_code, region_name, region_head_code, region_head_name, agency_name, agency_code, sub_channel_name, sub_channel_code, channel_name, channel_code, active, created_date, updated_date
)
select 
team_code, team_name, team_head_code, team_head_name, office_code, office_name, office_head_code, office_head_name, zone_code, zone_name, zone_head_code, zone_head_name, region_code, region_name, region_head_code, region_head_name, agency_name, agency_code, sub_channel_name, sub_channel_code, channel_name, channel_code, active, created_date, updated_date 
from raw_adlist_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_agentlist_profile
truncate table raw_agentlist_profile;
insert into raw_agentlist_profile (
mastercd, agentcd, agent_full_name, agent_first_name, applicationcd, application_date, joining_date, gender, title, date_of_birth, married_, education_level, course, major, institution, occupation, id_number, id_date, id_place, mobile_phone, home_phone, home_extension, business_phone, business_extension, email, res_address_province, res_address_district, res_address_ward, res_address_house_no_street, con_address_province, con_address_district, con_address_ward, con_address_house_no_street, tem_address_province, tem_address_district, tem_address_ward, tem_address_house_no_street, bank_name, bank_branch_name, bank_account, account_status, effective_date_from, pay_to, payment_mode, account, overridden_minimum_amount, taxid, mof_licence_no, mof_licence_eff_date_from, partfulltimecd, partfulltime_name, created_date, updated_date
) select
mastercd, agentcd, agent_full_name, agent_first_name, applicationcd, application_date, joining_date, gender, title, date_of_birth, married_, education_level, course, major, institution, occupation, id_number, id_date, id_place, mobile_phone, home_phone, home_extension, business_phone, business_extension, email, res_address_province, res_address_district, res_address_ward, res_address_house_no_street, con_address_province, con_address_district, con_address_ward, con_address_house_no_street, tem_address_province, tem_address_district, tem_address_ward, tem_address_house_no_street, bank_name, bank_branch_name, bank_account, account_status, effective_date_from, pay_to, payment_mode, account, overridden_minimum_amount, taxid, mof_licence_no, mof_licence_eff_date_from, partfulltimecd, partfulltime_name, created_date, updated_date
from raw_agentlist_profile_update
;

delete from raw_agentlist_hierarchy where business_date >= param_business_date; 
insert into raw_agentlist_hierarchy(
mastercd, agentcd, channelcd, channel_name, subchannelcd, subchannel_name, agencycd, agency_name, regioncd, region_name, regionheadcd, region_head_name, zonecd, zone_name, zone_head_cd, zone_head_name, teamcd, team_name, teamheadcd, team_head_name, officecd, office_name, officeheadcd, officeheadname, point_of_sale_cd, point_of_sale_name, bmcd, bm_name, branchcd, branch_name, reference_no, umcd, um_name, unitcd, unit_name, reportingcd, reporting_name, unitcd_of_reportingcd, unit_name_of_reportingcd, introducercd, introducer_name, recuitercd, recuiter_name, mentorcd, mentor_name, joining_date, approved_date, agent_designation, staffcd, agent_status, classification, termination_date, termination_reason, reinstatement_date, reinstatement_reason, created_date, updated_date, business_date
) select 
mastercd, agentcd, channelcd, channel_name, subchannelcd, subchannel_name, agencycd, agency_name, regioncd, region_name, regionheadcd, region_head_name, zonecd, zone_name, zone_head_cd, zone_head_name, teamcd, team_name, teamheadcd, team_head_name, officecd, office_name, officeheadcd, officeheadname, point_of_sale_cd, point_of_sale_name, bmcd, bm_name, branchcd, branch_name, reference_no, umcd, um_name, unitcd, unit_name, reportingcd, reporting_name, unitcd_of_reportingcd, unit_name_of_reportingcd, introducercd, introducer_name, recuitercd, recuiter_name, mentorcd, mentor_name, joining_date, approved_date, agent_designation, staffcd, agent_status, classification, termination_date, termination_reason, reinstatement_date, reinstatement_reason, created_date, updated_date, last_day(param_business_date) as business_date
from raw_agentlist_hierarchy_update;


-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_manpower_active_ratio
delete from raw_manpower_active_ratio where bussinessdate >= param_business_date;
insert into raw_manpower_active_ratio (
agent_code, agent_name, agent_designation, staffcd, classification, monthstart, monthend, active, activeex, activesp, activesp6m, ip, supervisor_code, supervisor_code_designation, unitcd, unit_name, branchcd, branch_name, teamcd, team_name, officecd, office_name, zonecd, zone_name, regioncd, region_name, channelcd, channel_name, joining_date, termination_date, reinstatement_date, bussinessdate, activefirstmonth, m1, created_date, updated_date
) select 
agent_code, agent_name, agent_designation, staffcd, classification, monthstart, monthend, active, activeex, activesp, activesp6m, ip, supervisor_code, supervisor_code_designation, unitcd, unit_name, branchcd, branch_name, teamcd, team_name, officecd, office_name, zonecd, zone_name, regioncd, region_name, channelcd, channel_name, joining_date, termination_date, reinstatement_date, bussinessdate, activefirstmonth, m1, created_date, updated_date
from raw_manpower_active_ratio_update
where bussinessdate >= param_business_date;

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# raw_kpitotal
delete from raw_kpitotal where bussinessdate >= date_format(param_business_date, '%Y%m%d');
insert into raw_kpitotal (
bussinessdate, duedate, rdocnum, productcode, premium, billingfrequency, 
`share`, txncode, policyyear, premiumterm, agcode, agcodedesignation, agent_status, 
classification, supervisor_code, supervisor_codedesignation, unitcd, unit, 
branchcd, branch, teamcd, team, officecd, office, zonecd, zone, regioncd, region, 
channelcd, channel, subchannelcd, subchannel_name, sacstyp, ip, ape, `case`, 
fyp, 2yp, 3yp, 4yp, 5yp, ryp, ultp, ic, fyc, 2yc, 3yc, 4yc, 5yc, ryc, comrate, 
com, ipflcan, icflcan, fypflcan, fycflcan, ccflcan, ulbnflcan, cibn, ulbn, ulbn1, 
sumins, tranno, life, hissdte, hoissdte, hprrcvdt, bankkey, countriders, 
countridersreal, cibnn, cibnn1, cibnnflcan, agent_idnumber, policyowner_idnumber, 
crosssales, rider
)
select 
bussinessdate, duedate, rdocnum, productcode, premium, billingfrequency, 
`share`, txncode, policyyear, premiumterm, agcode, agcodedesignation, agent_status, 
classification, supervisor_code, supervisor_codedesignation, unitcd, unit, 
branchcd, branch, teamcd, team, officecd, office, zonecd, zone, regioncd, region, 
channelcd, channel, subchannelcd, subchannel_name, sacstyp, ip, ape, `case`, 
fyp, 2yp, 3yp, 4yp, 5yp, ryp, ultp, ic, fyc, 2yc, 3yc, 4yc, 5yc, ryc, comrate, 
com, ipflcan, icflcan, fypflcan, fycflcan, ccflcan, ulbnflcan, cibn, ulbn, ulbn1, 
sumins, tranno, life, hissdte, hoissdte, hprrcvdt, bankkey, countriders, 
countridersreal, cibnn, cibnn1, cibnnflcan, agent_idnumber, policyowner_idnumber, 
crosssales, rider
from raw_kpitotal_update where bussinessdate >= date_format(param_business_date, '%Y%m%d');

-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
truncate table raw_gatotal;
insert into raw_gatotal (
region_name, officecd, office_name, gacode, ganame, agcode, agname, gatype, contractdt, effrom, effto, classification, housing, bonus12stend, q01, q02, q03, size, cashadv, targetlevel, openingdt, created_date, updated_date
)
select 
region_name, officecd, office_name, gacode, ganame, agcode, agname, gatype, contractdt, effrom, effto, classification, housing, bonus12stend, q01, q02, q03, size, cashadv, targetlevel, openingdt, created_date, updated_date
from raw_gatotal_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

truncate table raw_target;
insert into raw_target (
teamname, teamcode, ape, activeag, newrecruit, newleader, newagent, fyp, txndate, created_date, updated_date
)
select
teamname, teamcode, ape, activeag, newrecruit, newleader, newagent, fyp, txndate, created_date, updated_date
from raw_target_update;
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- ------last run----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
call job_segment(param_business_date, last_day(param_business_date));
-- ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

END















CREATE DEFINER=`root`@`localhost` PROCEDURE `get_agentlist`(
   IN from_date date, IN to_date date, tem_table varchar(255))
BEGIN


/***********************************
agent list of the month
*/
DROP TEMPORARY TABLE IF EXISTS re;

CREATE TEMPORARY TABLE re
AS
  SELECT 
	`ah`.`agentcd` AS `agentcd`,
    `ah`.`channelcd` AS `channelcd`,
	`ah`.`regioncd` AS `regioncd`,
	`ah`.`zonecd` AS `zonecd`,
	`ah`.`teamcd` AS `teamcd`,
	`ah`.`officecd` AS `officecd`,
	`ah`.`joining_date` AS `joining_date`,
	`ah`.`agent_designation` AS `agent_designation`,
	`ah`.`staffcd` AS `staffcd`,
	`ah`.`agent_status` AS `agent_status`,
	`ah`.`classification` AS `classification`,
	`ah`.`business_date` AS `business_date`,
     segment,
     last_day(ah.joining_date) AS m01,
	 last_day(ah.joining_date + INTERVAL 1 MONTH) m02,
	 last_day(ah.joining_date + INTERVAL 2 MONTH) m03,
	 last_day(ah.joining_date + INTERVAL 3 MONTH) m04,
	 last_day(ah.joining_date + INTERVAL 4 MONTH) m05,
	 last_day(ah.joining_date + INTERVAL 5 MONTH) m06,
	 last_day(ah.joining_date + INTERVAL 6 MONTH) m07,
	 last_day(ah.joining_date + INTERVAL 7 MONTH) m08,
	 last_day(ah.joining_date + INTERVAL 8 MONTH) m09,
	 last_day(ah.joining_date + INTERVAL 9 MONTH) m10,
	 last_day(ah.joining_date + INTERVAL 10 MONTH) m11,
	 last_day(ah.joining_date + INTERVAL 11 MONTH) m12,
	 last_day(ah.joining_date + INTERVAL 12 MONTH) m13,
	 last_day(ah.joining_date + INTERVAL 13 MONTH) m14
	FROM `raw_agentlist_hierarchy` `ah`
   WHERE     1 = 1
		 AND `ah`.`business_date` between from_date and to_date
		 AND lower(`ah`.`channelcd`) = 'tiedagency'
		 -- AND lower(`ah`.`agent_status`) <> 'terminated'
		 -- AND lower(`ah`.`regioncd`) not in (select region_code from raw_adlist where trim(lower(region_name)) = 'dummy')
         ;


/***********************************
return
*/
         
if tem_table is not null and tem_table <> '' then
	set @sql_com = concat('drop temporary table if exists ', tem_table);
	PREPARE stmt FROM @sql_com;
	EXECUTE stmt;
	set @sql_com = concat('create temporary table ', tem_table,' select * from re');
    
    PREPARE stmt FROM @sql_com;
	EXECUTE stmt;
else
	set @sql_com = 'select * from re';
    PREPARE stmt FROM @sql_com;
	EXECUTE stmt;
end if ;        

DEALLOCATE PREPARE stmt;
END


















CREATE DEFINER=`root`@`localhost` PROCEDURE `report_manpower_tiedagency`(
   IN from_date date, IN to_date date)
BEGIN

-- -------------------------------------------------------------------------------------------------------------------------------------------------
drop temporary table if exists tempfinalresult;
drop temporary table if exists tempsegoder;
create temporary table tempsegoder as
select seg.*, @idx:=@idx+1 as idx from (
select 'GEN LION (from Apr ''17)' as segment union
select 'Rookie in month' as segment union
select 'Rookie last month' as segment union
select '2-3 months' as segment union
select '4-6 mths' as segment union
select '7-12mth' as segment union
select '13+mth' as segment union
select 'SA' as segment union
select 'Total (excl. SA)' as segment union
select 'Total' as segment 

) seg, (select @idx:=0) t;
-- -------------------------------------------------------------------------------------------------------------------------------------------------
call get_agentlist(from_date, to_date,'tempagentlist');
delete from tempagentlist where lower(agent_status) = 'terminated' or agentcd is null or agentcd = '';
-- -------------------------------------------------------------------------------------------------------------------------------------------------
drop temporary table if exists tempdatatamtonghop;
create temporary table tempdatatamtonghop as
select 
business_date, agentcd, channelcd, regioncd, zonecd, officecd, teamcd, segment
from 
tempagentlist
;
-- -------------------------------------------------------------------------------------------------------------------------------------------------


drop temporary table if exists result_table;
create temporary table result_table as
-- -------------------------------------------------------------------------------------------------------------------------------------------------
select business_date, segment, count(agentcd) as kpivalue 
from tempdatatamtonghop
group by business_date, segment
;
-- -------------------------------------------------------------------------------------------------------------------------------------------------
alter table result_table change column segment segment varchar(50);
-- -------------------------------------------------------------------------------------------------------------------------------------------------
insert into result_table
select business_date, 'Total' as segment, count(agentcd) as kpivalue  
from tempdatatamtonghop
group by business_date

;-- -------------------------------------------------------------------------------------------------------------------------------------------------
-- -------------------------------------------------------------------------------------------------------------------------------------------------
insert into result_table
select business_date, 'Total (excl. SA)' as segment, count(agentcd) as kpivalue 
from tempdatatamtonghop
where segment != 'SA'
group by business_date

;-- -------------------------------------------------------------------------------------------------------------------------------------------------

-- -------------------------------------------------------------------------------------------------------------------------------------------------
# pivot result
set @sql1 = (
	select group_concat(
		distinct concat('sum(case when business_date=''', business_date, ''' then kpivalue else 0 end) as ', '`', date_format(business_date, '%Y%m'),'`') order by business_date
	) from result_table
);

set @sql2 = concat(
	' create temporary table tempfinalresult as ',
	' select segment, ', @sql1,
    ' from result_table  ',
    ' group by segment '
);
    
PREPARE stmt FROM @sql2;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;


select a.* from tempfinalresult a 
left join tempsegoder b on a.segment = b.segment
order by b.idx
;


END