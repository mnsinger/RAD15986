import ibm_db
import getpass
import datetime
import os
import mskcc
import shutil
import xlrd
import xlsxwriter
import pypyodbc
import sys
import re

print(datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))

###########################
#       CONNECTION        #
###########################

input_file_1 = '../properties.txt'
f_in = open(input_file_1, 'r')
properties_dict = {}
for line in f_in:
    properties_dict[line.partition('=')[0]] = line.partition('=')[2].strip()
f_in.close()

connection_idb = ibm_db.connect('DATABASE=DB2P_MF;'
                     'HOSTNAME=ibm3270;'
                     'PORT=3021;'
                     'PROTOCOL=TCPIP;'
                     'UID={0};'.format(properties_dict["idb_service_uid1"]) +
                     'PWD={0};'.format(mskcc.decrypt(properties_dict["idb_service_pwd1"]).decode("latin-1")), '', '')

connection_darwin = ibm_db.connect('DATABASE=DVPDB01;'
                     'HOSTNAME=pidvudb1;'
                     'PORT=51013;'
                     'PROTOCOL=TCPIP;'
                     'UID={0};'.format(properties_dict["darwin_uid"]) +
                     'PWD={0};'.format(mskcc.decrypt(properties_dict["darwin_pwd"]).decode("latin-1")), '', '')

connection_sql_server = pypyodbc.connect("Driver={{SQL Server}};Server={};Database={};Uid={};Pwd={};".format(
                    "PS23A,61692",
                    "DEDGPDLR2D2",
                    properties_dict["sqlserver_ps23a_uid"],
                    mskcc.decrypt(properties_dict["sqlserver_ps23a_pwd"]).decode("latin-1")
                    )
                )

###########################
#         DECLARE         #
###########################

now_raw = datetime.datetime.now()
now = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
dataline_report_number = "RAD15986" #os.path.basename(__file__).replace(".py", "")

# file vars
input_file_1 = r"\\pens62\sddshared\Clinical Systems\Data Administration\DataLine\Plans\Radiology\RAD15986\MRN Scan Date.xlsx"
input_file_2 = r"\\pens62\sddshared\Clinical Systems\Data Administration\DataLine\Plans\Radiology\RAD15986\Lab.xlsx"
output_file_1 = '{}-{}.log'.format(dataline_report_number, now)
output_file_2 = '{}-progress-{}.log'.format(dataline_report_number, now)

# excel vars
excel_file_name = '{}-{}.xlsx'.format(dataline_report_number, now)
workbook = xlsxwriter.Workbook(excel_file_name)
worksheet_dx = workbook.add_worksheet('Diagnoses')
worksheet_staging = workbook.add_worksheet('Staging')
worksheet_labs = workbook.add_worksheet('Labs')
worksheet_rt = workbook.add_worksheet('Prior RT')
worksheet_surgery = workbook.add_worksheet('Surgery')
worksheet_chemo = workbook.add_worksheet('Chemo')
worksheet_bp = workbook.add_worksheet('Blood Pressure')
worksheet_hml = workbook.add_worksheet('Home Med')
worksheet_dx_codes = workbook.add_worksheet('Dx Codes')
worksheet_criteria = workbook.add_worksheet('Criteria')

###########################
#         CLASSES         #
###########################

class Patient:
    'Common base class for all patients'
    patient_count = 0

    def __init__(self):
        self.deid = None
        self.scan_dte = None
        self.lab_dict = {} # name -> data
        self.bp_tup = {}
        self.bp_tup["sys"] = None
        self.bp_tup["dias"] = None

class Lab:
    def __init__(self):
        self.lab_date = None
        self.lab_value = None
        self.lab_days_from_scan = None

    def to_string(self):
        return "lab_date: {}, lab_value: {}, lab_days_from_scan: {}".format(self.lab_date, self.lab_value, self.lab_days_from_scan)

###########################
#        FUNCTIONS        #
###########################

def output_excel_dataline_info(worksheet):

  SQL = """
      SELECT   Criteria,
               [Project Description],
               [Data Elements]
      FROM     [DEDGPDLR2D2].dbo.projects
      WHERE    [Project Code] = '{}'
  """.format(dataline_report_number)

  cursor = connection_sql_server.cursor()
  cursor.execute(SQL)

  row = {}
  row_raw = cursor.fetchone()
  if row_raw is not None:
      columns = [column[0] for column in cursor.description]
      row = row_to_dict(row_raw, columns)

  
  criteria_text = "{}\n{}".format(row["data elements"], row["criteria"])
  options = {
      'width': 1000,
      'height': 800,
  }
  worksheet_criteria.insert_textbox(5, 2, criteria_text, options)

  fmt = workbook.add_format({'bold': True})
  worksheet.write(0, 0, row["project description"], fmt)
  worksheet.write(1, 0, "DataLine Report {}".format(dataline_report_number), fmt)
  worksheet.write(2, 0, "Produced on {}, by DataLine in Information Systems".format(now_raw.strftime('%B %d, %Y')))
  worksheet.write(3, 0, 'See "Criteria" sheet for inclusion criteria')

  cursor.close()

def output_excel_column_headers(worksheet, in_stmt, col_start, row):
  fmt = workbook.add_format({'bold': True, 'bg_color': '#C5D9F1'})
  d=col_start
  for n in range(0, ibm_db.num_fields(in_stmt)):
      field_width = len(ibm_db.field_name(in_stmt, n))+3
      if 'MRN' in ibm_db.field_name(in_stmt, n):
          field_width = 10
      worksheet.write(row, d+n, ibm_db.field_name(in_stmt, n), fmt)
      worksheet.set_column(d+n, d+n, field_width)

def output_excel_rows(worksheet, stmt, in_dict, row):
  while in_dict != False:
      for key in in_dict.keys():
          if isinstance(key, int) and in_dict[key] != 'None':
              if isinstance(in_dict[key], datetime.date):
                  worksheet.write(row, key, in_dict[key].strftime('%Y-%m-%d'))
              elif isinstance(in_dict[key], str):
                  worksheet.write(row, key, in_dict[key].strip())
              else:
                  worksheet.write(row, key, in_dict[key])
              
      in_dict = ibm_db.fetch_both(stmt)
      row += 1
  return row

def db2_row_to_list(in_dict):
  row = []
  for key in in_dict.keys():
    if isinstance(key, int) and in_dict[key] != 'None':
      row.append(in_dict[key])
  return row

def output_excel_list(worksheet, in_list, row):
  c=0
  for cell in in_list:
    if isinstance(cell, datetime.date):
      worksheet.write(row, c, cell.strftime('%Y-%m-%d'))
    elif isinstance(cell, str):
      worksheet.write(row, c, cell.strip())
    else:
      worksheet.write(row, c, cell)
    c+=1
  return 0

def output_excel_header_list(worksheet, in_list, row):
  c=0
  fmt = workbook.add_format({'bold': True, 'bg_color': '#C5D9F1'})
  for cell in in_list:
    if isinstance(cell, datetime.date):
      worksheet.write(row, c, cell.strftime('%Y-%m-%d'), fmt)
    elif isinstance(cell, str):
      worksheet.write(row, c, cell.strip(), fmt)
    else:
      worksheet.write(row, c, cell, fmt)
    c+=1
  return 0

def row_to_dict(row_raw, columns):
  row = {}
  x = 0
  for col in columns:
      row[col] = row_raw[x]
      x += 1
  return row

###########################
#          MAIN           #
###########################

with open(output_file_1, 'a') as f:
    f.write("{}\n".format(now))

output_excel_dataline_info(worksheet_dx)

# READ DATA FROM EXCEL

book = xlrd.open_workbook(input_file_1)
sheet = book.sheet_by_index(0)

mrns = {}
lab_names = set()

# MRNS AND SCAN DATES FROM EXCEL

for row_index in range(1, sheet.nrows):
  mrn = str(int(sheet.cell(row_index, 0).value)).strip().zfill(8)
  py_date1 = sheet.cell(row_index, 1).value #datetime.date(year, month, day)
  py_date1 = datetime.datetime.strptime(py_date1, '%Y-%m-%d').date()
  print(py_date1)

  mrns[mrn] = Patient()
  mrns[mrn].scan_dte = py_date1

mrn_list_string = (', '.join("'" + item + "'" for item in mrns.keys()))

# LAB SUBTEST NAMES FROM EXCEL

book = xlrd.open_workbook(input_file_2)

for x in range(0,5):
  sheet = book.sheet_by_index(x)

  for row_index in range(1, sheet.nrows):
    subtest_name = sheet.cell(row_index, 1).value.strip()
    lab_names.add(subtest_name)

lab_list_string = (', '.join("'" + item + "'" for item in lab_names))

# GET MRNS -> DEIDS AND CREATE A LOOKUP DICTIONARY

sql_string = """
            select trim(PT_MRN) PT_MRN, PT_PT_DEIDENTIFICATION_ID DEID
            from dv.patient_demographics
            where pt_mrn in ({})
""".format(mrn_list_string, lab_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_darwin, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

while db_dict != False:
  mrn = db_dict["PT_MRN"]
  mrns[mrn].deid = db_dict["DEID"]
  db_dict = ibm_db.fetch_assoc(stmt)

# GET DIAGNOSIS DATA

sql_string = """
        SELECT trim(TM_MRN) MRN,
               AV_DESC AS CASE_STS,
               TM_DX_DTE,
               TM_HIST_CD, 
               H.CLM_CLSF_DESC_MSK AS HIST_DESC,
               TM_SITE_CD,
               S.CLM_CLSF_DESC_MSK AS SITE_DESC
          FROM IDB.CDB_TUMOR
          JOIN IDB.CLM H ON H.CLM_CLSF_CD = TM_HIST_CD
          JOIN IDB.CLM S ON S.CLM_CLSF_CD = TM_SITE_CD 
          JOIN IDB.ALLOWVALS
            ON AV_ELEMENT = 'CASE_STS'
           AND AV_CODE = TM_CASE_STS
           AND TM_CASE_STS IN ('1', '3', '6', '7', '8')
           AND TM_MRN in ({})
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_dx, ["MRN", "DEID", "CASE_STS", "TM_DX_DTE", "TM_HIST_CD", "HIST_DESC", "TM_SITE_CD", "SITE_DESC"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  output_excel_list(worksheet_dx, [mrn, mrns[mrn].deid, db_dict["CASE_STS"], db_dict["TM_DX_DTE"], db_dict["TM_HIST_CD"], db_dict["HIST_DESC"], db_dict["TM_SITE_CD"], db_dict["SITE_DESC"]], row)
  row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_dx, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1


sql_string = """
        SELECT   trim(tm_mrn) as "MRN", 
                 TM_CLIN_TNM_T as "Clinical T Stage",
                 TM_CLIN_TNM_N as "Clinical N Stage",
                 TM_CLIN_TNM_M as "Clinical M Stage",
                 TM_CLIN_STG_GRP as "Clinical Group Stage",
                 TM_PATH_TNM_T as "Pathologic T Stage",
                 TM_PATH_TNM_N as "Pathologic N Stage",
                 TM_PATH_TNM_M as "Pathologic M Stage",
                 TM_PATH_STG_GRP as "Pathologic Group Stage"
        FROM     idb.cdb_tumor
        WHERE    TM_CASE_STS IN ('1', '3', '6', '7', '8')
        AND TM_MRN in ({})
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_staging, ["MRN", "DEID", "Clinical T Stage", "Clinical N Stage", "Clinical M Stage", "Clinical Group Stage", "Pathologic T Stage", "Pathologic N Stage", "Pathologic M Stage", "Pathologic Group Stage"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  output_excel_list(worksheet_staging, [mrn, mrns[mrn].deid, db_dict["Clinical T Stage"], db_dict["Clinical N Stage"], db_dict["Clinical M Stage"], db_dict["Clinical Group Stage"], db_dict["Pathologic T Stage"], db_dict["Pathologic N Stage"], db_dict["Pathologic M Stage"], db_dict["Pathologic Group Stage"]], row)
  row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_staging, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1


sql_string = """
            select trim(RTC_TX_MRN) MRN,
                   RTC_TX_COURSE_NO,
                   RTC_TX_PLANSETUP_NO,
                   RTC_TX_COURSE_ID,
                   RTC_TX_PLAN_NAME, --Tells you the site
                   RTC_TX_PLANNED_FRACTIONS,
                   RTC_TX_DELIVERED_FRACTIONS,
                   RTC_TX_PLANNED_DOSE, 
                   RTC_TX_DELIVERED_DOSE,
                   RTC_TX_START_DTE,
                   RTC_TX_STOP_DTE,
                   RTC_TX_ELAPSED_DAYS,
                   RTC_PRIMARY_REF_POINT, -- Sometimes less technical or less granular version of the site than in RTC_TX_PLAN_NAME
                   RTC_DOSE_CORRECTION
              from idb.radonc_treatment_course
              where rtc_tx_mrn in ({})
              order by RTC_TX_MRN, RTC_TX_START_DTE
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_rt, ["RTC_TX_MRN","DEID","SCAN_DATE","RTC_TX_COURSE_NO","RTC_TX_PLANSETUP_NO","RTC_TX_COURSE_ID","RTC_TX_PLAN_NAME","RTC_TX_PLANNED_FRACTIONS","RTC_TX_DELIVERED_FRACTIONS","RTC_TX_PLANNED_DOSE","RTC_TX_DELIVERED_DOSE","RTC_TX_START_DTE","RTC_TX_STOP_DTE","RTC_TX_ELAPSED_DAYS","RTC_PRIMARY_REF_POINT","RTC_DOSE_CORRECTION"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  if db_dict["RTC_TX_START_DTE"] < mrns[mrn].scan_dte:
    output_excel_list(worksheet_rt, [mrn, mrns[mrn].deid, mrns[mrn].scan_dte, db_dict["RTC_TX_COURSE_NO"],db_dict["RTC_TX_PLANSETUP_NO"],db_dict["RTC_TX_COURSE_ID"],db_dict["RTC_TX_PLAN_NAME"],db_dict["RTC_TX_PLANNED_FRACTIONS"],db_dict["RTC_TX_DELIVERED_FRACTIONS"],db_dict["RTC_TX_PLANNED_DOSE"],db_dict["RTC_TX_DELIVERED_DOSE"],db_dict["RTC_TX_START_DTE"],db_dict["RTC_TX_STOP_DTE"],db_dict["RTC_TX_ELAPSED_DAYS"],db_dict["RTC_PRIMARY_REF_POINT"],db_dict["RTC_DOSE_CORRECTION"]], row)
    row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_rt, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1




sql_string = """
              select trim(CPO_MRN) MRN, CPO_ORD_NAME, CPO_START_DTE
              from idb.chemo_performed_orders
              where CPO_MRN in ({})
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_chemo, ["MRN","DEID","SCAN DATE","ORDER NAME","START DATE"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  if db_dict["CPO_START_DTE"] < mrns[mrn].scan_dte:
    output_excel_list(worksheet_chemo, [mrn, mrns[mrn].deid, mrns[mrn].scan_dte, db_dict["CPO_ORD_NAME"],db_dict["CPO_START_DTE"]], row)
    row += 1
    if mrn in unique_mrns:
      del unique_mrns[mrn]

  db_dict = ibm_db.fetch_assoc(stmt)

for mrn in unique_mrns:
  output_excel_list(worksheet_chemo, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1





sql_string = """
            SELECT trim(SSE_MRN) MRN,
                   SSE_SURG_DTE,
                   SSE_LOG_STS,
                   SSP_PROC_CPT4_CD,
                   SSP_PROC_CPT4_DESC,
                   SSP_SURG_LAST_NM,
                   SSP_SURG_FIRST_NM,
                   SSP_SURG_SVC_CD
              FROM IDB.SRG_SURG_EVENT 
              JOIN IDB.SRG_SURG_PROCEDURE ON SSE_LOG_ID = SSP_LOG_ID
              where SSE_MRN in ({})
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_surgery, ["MRN","DEID","SCAN DATES","SSE_SURG_DTE","SSE_LOG_STS","SSP_PROC_CPT4_CD","SSP_PROC_CPT4_DESC","SSP_SURG_LAST_NM","SSP_SURG_FIRST_NM","SSP_SURG_SVC_CD"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  if db_dict["SSE_SURG_DTE"] < mrns[mrn].scan_dte:
    output_excel_list(worksheet_surgery, [mrn, mrns[mrn].deid, mrns[mrn].scan_dte, db_dict["SSE_SURG_DTE"],db_dict["SSE_LOG_STS"],db_dict["SSP_PROC_CPT4_CD"],db_dict["SSP_PROC_CPT4_DESC"],db_dict["SSP_SURG_LAST_NM"],db_dict["SSP_SURG_FIRST_NM"],db_dict["SSP_SURG_SVC_CD"]
], row)
    row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_surgery, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1


sql_string = """
select trim(hcpr_mrn) mrn, hcpr_created_dte start_dte, trim(HCPR_COMMENTS) comments, trim(HCPR_DRUG_NAME) drug_name
from IDB.HML_CLIENT_PRESCRIPTION
where hcpr_mrn in ({})
""".format(mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_hml, ["MRN","DEID","SCAN DATES","START DATE","DRUG NAME","COMMENTS"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  if db_dict["START_DTE"] < mrns[mrn].scan_dte:
    output_excel_list(worksheet_hml, [mrn, mrns[mrn].deid, mrns[mrn].scan_dte, db_dict["START_DTE"],db_dict["DRUG_NAME"],db_dict["COMMENTS"]], row)
    row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_hml, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1







# I25.1: coronary artery disease
# E03: hypothyroidism
# E11: Type II diabetes
# I10: hypertension
# E78.2, E78.4, and E78.5: hyperlipidemia
# E88.81: metabolic syndrome
# K76.0: non-alcoholic fatty liver disease
# R73.03: pre-diabetes

sql_string = """
SELECT trim(CC_MRN) MRN, 
       MIN(CC_EFF_DTE) AS MIN_ICD_EFF_DTE,
       CC_CLSF_CD AS ICD_CD,
       CLM_CLSF_DESC_MSK AS ICD_DESC
  FROM IDB.CTC_CLSF
  JOIN IDB.CLM
    ON CC_CLSF_CD = CLM_CLSF_CD
 WHERE (
        CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}'
        OR
        CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}' OR CC_CLSF_CD LIKE '{}'
        )
   AND CC_CLSF_TYPE LIKE 'DF%' --FINAL DIAGNOSIS
   AND (CC_PT_TYPE = 'I' --INPATIENT
    OR (CC_PT_TYPE = 'O' AND CC_FINALIZED_IND = 'Y')) --CIC-CODED OUTPATIENT
   AND CC_MRN in ({})
 GROUP
    BY CC_MRN, 
       CC_CLSF_CD,
       CLM_CLSF_DESC_MSK
""".format('I25.1%','E03%','E11%', 'I10%', 'E78.2%', 'E78.4%', 'E78.5%', 'E88.81%', 'K76.0%', 'R73.03%', '401%', '414%', '224.9%', '250%', '272%', '277.7%', '571.8%', '790.29%', mrn_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

unique_mrns = mrns.copy()
output_excel_header_list(worksheet_dx_codes, ["MRN","DEID","SCAN DATE","DX DATE","CODE","DESC"], 6)
row=7
while db_dict != False:
  mrn = db_dict["MRN"]
  output_excel_list(worksheet_dx_codes, [mrn, mrns[mrn].deid, mrns[mrn].scan_dte, db_dict["MIN_ICD_EFF_DTE"],db_dict["ICD_CD"],db_dict["ICD_DESC"]], row)
  row += 1
  db_dict = ibm_db.fetch_assoc(stmt)
  if mrn in unique_mrns:
    del unique_mrns[mrn]

for mrn in unique_mrns:
  output_excel_list(worksheet_dx_codes, [mrn, mrns[mrn].deid, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], row)
  row += 1














sql_string = """
            select trim(LR_MRN) LR_MRN, trim(LR_TEST_NAME) LR_TEST_NAME, trim(LR_SUBTEST_NAME) LR_SUBTEST_NAME, ' ' AS SCAN_DATE, LR_PERFORMED_DTE , trim(LR_RESULT_VALUE) LR_RESULT_VALUE, LR_TEST_LOW_LIMIT , LR_TEST_UP_LIMIT 
            from idb.lab_results
            where lr_mrn in ({}) and lr_subtest_name in ({}) and lr_result_value <> ' '
""".format(mrn_list_string, lab_list_string)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

while db_dict != False:
  mrn = db_dict["LR_MRN"]
  scan_dte = mrns[mrn].scan_dte
  if db_dict["LR_SUBTEST_NAME"] not in mrns[mrn].lab_dict:
    lab = Lab()
    lab.lab_test = db_dict["LR_TEST_NAME"]
    lab.lab_low = db_dict["LR_TEST_LOW_LIMIT"]
    lab.lab_high = db_dict["LR_TEST_UP_LIMIT"]
    lab.lab_date = db_dict["LR_PERFORMED_DTE"]
    lab.lab_value = db_dict["LR_RESULT_VALUE"]
    lab.lab_days_from_scan = abs(db_dict["LR_PERFORMED_DTE"] - scan_dte)

    #print(lab.to_string())

    mrns[mrn].lab_dict[db_dict["LR_SUBTEST_NAME"]] = lab

  else:
    existing_lab = mrns[mrn].lab_dict[db_dict["LR_SUBTEST_NAME"]]
    if abs(db_dict["LR_PERFORMED_DTE"] - scan_dte) < existing_lab.lab_days_from_scan:
      lab = Lab()
      lab.lab_test = db_dict["LR_TEST_NAME"]
      lab.lab_low = db_dict["LR_TEST_LOW_LIMIT"]
      lab.lab_high = db_dict["LR_TEST_UP_LIMIT"]
      lab.lab_date = db_dict["LR_PERFORMED_DTE"]
      lab.lab_value = db_dict["LR_RESULT_VALUE"]
      lab.lab_days_from_scan = abs(db_dict["LR_PERFORMED_DTE"] - scan_dte)

      #print(lab.to_string())

      mrns[mrn].lab_dict[db_dict["LR_SUBTEST_NAME"]] = lab

  db_dict = ibm_db.fetch_assoc(stmt)

blood_pressure_items = """'Blood Pressure BP Systolic', 'Blood Pressure Diastolic', 'Blood Pressure Systolic' """ 

sql_string = """
        SELECT 
               trim(CDD_MRN) MRN, CASE WHEN CDO_ITEM_NAME LIKE '%Systolic%' THEN 'sys' ELSE 'dias' END bp_type,
               CDD_AUTHORED_DT,
               CDO_ITEM_NAME,
               CDO_VALUE_TEXT
          FROM IDB.CD_DOCUMENT
          JOIN IDB.CD_OBSERVATION ON CDD_DOC_GUID = CDO_DOC_GUID AND CDD_AUTHORED_DTE = CDO_AUTHORED_DTE AND CDD_CANCEL_IND <> '1' AND CDD_INCOMPLETE_IND <> '1' AND  CDD_MRN in ({}) AND CDO_ITEM_NAME in ({})
""".format(mrn_list_string, blood_pressure_items)

print(sql_string)

stmt = ibm_db.prepare(connection_idb, sql_string)

ibm_db.execute(stmt)

db_dict = ibm_db.fetch_assoc(stmt)

while db_dict != False:

  mrn = db_dict["MRN"]
  bp_type = db_dict["BP_TYPE"]
  scan_dte = mrns[mrn].scan_dte
  if not mrns[mrn].bp_tup[bp_type]:
    bp_date = db_dict["CDD_AUTHORED_DT"]
    bp_item = db_dict["CDO_ITEM_NAME"]
    bp_value = db_dict["CDO_VALUE_TEXT"]
    bp_days_from_scan = abs(db_dict["CDD_AUTHORED_DT"].date() - scan_dte)

    #print(lab.to_string())

    mrns[mrn].bp_tup[bp_type] = (bp_date, bp_item, bp_value, bp_days_from_scan)

  else:
    bp_date, bp_item, bp_value, bp_days_from_scan = mrns[mrn].bp_tup[bp_type]
    if abs(db_dict["CDD_AUTHORED_DT"].date() - scan_dte) < bp_days_from_scan:
      bp_date = db_dict["CDD_AUTHORED_DT"]
      bp_item = db_dict["CDO_ITEM_NAME"]
      bp_value = db_dict["CDO_VALUE_TEXT"]
      bp_days_from_scan = abs(db_dict["CDD_AUTHORED_DT"].date() - scan_dte)

      #print(lab.to_string())

      mrns[mrn].bp_tup[bp_type] = (bp_date, bp_item, bp_value, bp_days_from_scan)

  db_dict = ibm_db.fetch_assoc(stmt)

output_excel_header_list(worksheet_labs, ["MRN", "DEID", "SCAN DATE", "TEST NAME", "SUBTEST NAME", "PREFORMED DTE", "RESULT VALUE", "LOW LIMIT", "HIGH LIMIT"], 6)
output_excel_header_list(worksheet_bp, ["MRN", "DEID", "BP TYPE", "SCAN DATE", "AUTHORED DATE", "ITEM NAME", "VALUE"], 6)
lab_row=7
bp_row=7
for mrn in mrns:
  deid = mrns[mrn].deid
  if mrns[mrn].lab_dict:
    for lab_subtest in mrns[mrn].lab_dict:
      lab = mrns[mrn].lab_dict[lab_subtest]
      output_excel_list(worksheet_labs, [mrn, deid, mrns[mrn].scan_dte, lab.lab_test, lab_subtest, lab.lab_date, lab.lab_value, lab.lab_low, lab.lab_high], lab_row)
      lab_row += 1
  else:
    output_excel_list(worksheet_labs, [mrn, deid, mrns[mrn].scan_dte, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"], lab_row)
    lab_row += 1

  for bp_type in ['sys', 'dias']:
    if mrns[mrn].bp_tup[bp_type]:
      output_excel_list(worksheet_bp, [mrn, deid, bp_type, mrns[mrn].scan_dte, mrns[mrn].bp_tup[bp_type][0], mrns[mrn].bp_tup[bp_type][1], mrns[mrn].bp_tup[bp_type][2]], bp_row)
    else:
      output_excel_list(worksheet_bp, [mrn, deid, bp_type, mrns[mrn].scan_dte, "N/A", "N/A", "N/A"], bp_row)
    bp_row += 1

workbook.close()

shutil.move(excel_file_name, r"\\pens62\sddshared\Clinical Systems\Data Administration\DataLine\Plans\Radiology\{}\{}".format(dataline_report_number, excel_file_name))

print(datetime.datetime.now().strftime('%Y%m%d-%H%M%S'))