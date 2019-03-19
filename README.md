# RAD15986

## labs_dx_and_meds.py

This file (**labs_dx_and_meds.py**) is relatively neat and didn't have multiple changes and hacks from the initial spec.

The goal is to find various values that were closest to the scan date for a list of patients and scan dates.

- Reads in 11,000+ MRNs and scan dates and a few spreadheets of lab subtest names (~40) from an excel file. 
- Gets all MRNs and DEIDs and makes a dictionary / lookup (so you won't need to use SQL to get the DEIDs - they're local in memory).
- Each Patient is an object with the following attributes:
	* deid - deidentified ID (one per patient)
	* scan_date - date patient had radiology scan (one per patient)
	* lab_dict - dictionary of lab tests where lab test closest to scan date is kept
	* bp_tup 
		- blood pressure dictionary of 'sys' or 'dias' (systolic, diastolic). 
		- hack from when inital spec changed / was misunderstood. 
		- initially thought there was only one bp measure per patient.
		- stores (bp_date, bp_item, bp_value, bp_days_from_scan)
- Each Lab is an object stored in the patient.lab_dict attribute with the following attributes:
	* lab_date - date lab subtest occurred
	* lab_value - result of lab subtest
	* lab_days_from_scan - how many days from patient_scan date the lab occurred
- Gets Primary Cancer Diagnoses info and Cancer Staging info for all MRNs and outputs to Excel sheets. For any MRNs missing this info, output N/A to the Excel sheet.
- Gets Radiation Oncology Treatment Course (radonc) info for all MRNs and output to Excel sheet. For any MRNs missing this info, output N/A to the Excel sheet.
- Gets Chemo Performed Orders info for all MRNs and outputs to Excel sheet. For any MRNs missing this info, output N/A to the Excel sheet.
- Gets Surgery info for all MRNs and outputs to Excel sheet. For any MRNs missing this info, output N/A to the Excel sheet.
- Gets Client Prescription (home meds) info for all MRNs and outputs to Excel sheet. For any MRNs missing this info, output N/A to the Excel sheet.
- Gets all lab and blood pressure results for all MRNs and builds the Patient and Lab objects with the values closest to the scan date.
- Loops through the list of MRNs and outputs lab and blood pressure results to respective Excel sheets. For any MRNs missing info for a lab or blood pressure result, output N/A to the Excel sheet.
- Move the Excel file to the network drive
