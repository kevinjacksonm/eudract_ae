
from lxml import etree
import argparse
import os
import pandas as pd

# names of spreadsheets created by SP, one for serious events and one for non-serious events
nsae_xlsx="t-teae-nonser-5pct.xlsx"
sae_xlsx ="t-teae-ser.xlsx"

# set up XML namespaces necessary EudraCT XML file

ns_eudra_ae="http://eudract.ema.europa.eu/schema/clinical_trial_result/adverse_events"
ns_xsi="http://www.w3.org/2001/XMLSchema-instance"

# create dictionary of XML namespace mappings
NSMAP = { "ae"   : ns_eudra_ae,
					"xsi"   : ns_xsi}

# Suppress pandas warning						
pd.options.mode.chained_assignment = None

# Read Body System Code lookup from an outside spreadsheet into dataframe
df_spor_lookup = pd.read_excel("SPOR_lookup.xlsx", sheet_name="Terms")
spor_id=dict(zip(df_spor_lookup.Body_System, df_spor_lookup.SPOR_identifier))
spor_version=df_spor_lookup.Version[0]
spor_term_version=dict(zip(df_spor_lookup.Body_System, df_spor_lookup.Version))
		
###################################################################
# Parse and validate command line arguments
###################################################################
parser = argparse.ArgumentParser(description='Eudra CT utility ')
parser.add_argument('-d','--directory',		help='Directory path where Eudra CT spreadsheets are located')
parser.add_argument('-f','--file',			  help='Name of adverse events XML output file (without XML extension)')

args = parser.parse_args()

filename="eudract_adverse_events.xml"
if (args.file):
		filename=args.file+".xml"

if (args.directory):
		directory=args.directory
		if not directory.endswith('/'):
				directory=args.directory+'/';		
		if not (os.path.exists(directory)):
				print ('Directory ' + directory + 'specified in -d parameter does not exist')
				exit(1)
		else:
				if not os.path.isfile(directory+nsae_xlsx) or not os.path.isfile(directory+sae_xlsx):
						print ("Cannot locate one or both of spreadsheets", nsae_xlsx, 'or', sae_xlsx, 'in directory', directory)
						exit(1)
		filename=directory+filename
		print (filename)
else:
		print ('Directory ' + directory + ' must specified in -d parameter')
		exit(1)

###################################################################
# Write XML out to file
#
# root - root node of XML
###################################################################
def write_xml(root):

		# doing this makes pretty_print work
		for element in root.iter():
		    element.tail = None
			
		# create the xml string
		obj_xml = etree.tostring(root,
		                         pretty_print=True,
		                         xml_declaration=True, 
		                         encoding='UTF-8')
		
		
		with open(filename, "wb") as xml_writer:
				xml_writer.write(obj_xml)

###################################################################
# Build reporting groups section of XML file
#
# df_counts = dataframe loaded from spreadsheet with overall counts
###################################################################
def build_xml_reporting_groups(node, df_counts):
    repgroups_node=etree.SubElement(node,"reportingGroups")

    for index, row in df_counts.iterrows():
        repgroupnode=etree.SubElement(repgroups_node,"reportingGroup")
        
        repgroupnode.set("id","ReportingGroup-"+str(index))
        
        title=etree.SubElement(repgroupnode,"title")
       
        rpname=row["Treatment_Groups"]
        if len(rpname)>1:
        		title.text=rpname        	
        else:
        		title.text='Group ' + rpname        		

        description=etree.SubElement(repgroupnode,"description")
        description.text="Patients who received "+rpname
        subs_affected_nsae=etree.SubElement(repgroupnode,"subjectsAffectedByNonSeriousAdverseEvents")
        subs_affected_nsae.text=str(int(row["NSAE_Subjects_Affected"]))    
        subs_affected_sae=etree.SubElement(repgroupnode,"subjectsAffectedBySeriousAdverseEvents")
        subs_affected_sae.text=str(int(row["SAE_Subjects_Affected"]))
        subs_exposed=etree.SubElement(repgroupnode,"subjectsExposed")
        subs_exposed.text=str(int(row["Subjects_Exposed_Number"]))
        deaths_all_causes=etree.SubElement(repgroupnode,"deathsAllCauses")
        if not pd.isna(row["Fatalities_Number"]):        					
        		deaths_all_causes.text=str(int(row["Fatalities_Number"]))
        else:
            deaths_all_causes.text="0"            
        deaths_resulting_from_aes=etree.SubElement(repgroupnode,"deathsResultingFromAdverseEvents")
        if not pd.isna(row["Fatalities_Causally_Related_to_Treatment_Number"]):        
            deaths_resulting_from_aes.text=str(int(row["Fatalities_Causally_Related_to_Treatment_Number"]))
        else:
            deaths_resulting_from_aes.text="0"

###################################################################
# Build non-serious adverse events XML
#
# node = parent node with tag <ae:adverseEvents>
# df_terms = dataframe loaded from non-serious ae spreadsheet
# reporting_groups = dictionary of reportting groups with ID  #
###################################################################
def build_xml_non_serious_events(node, df_terms, reporting_groups):
    
		# non-serious adverse events
		nsaes=etree.SubElement(node,"nonSeriousAdverseEvents")
		organ_system=""
		pref_term=""
		for index, row in df_terms.iterrows():
				if organ_system != row["System_Organ_Class"] or pref_term != row["Preferred_Term"]:
						organ_system = row["System_Organ_Class"]
						pref_term = row["Preferred_Term"]
						nsae=etree.SubElement(nsaes,"nonSeriousAdverseEvent")
						term=etree.SubElement(nsae,"term")
						term.text=pref_term
						soc=etree.SubElement(nsae,"organSystem")
						eutctid=etree.SubElement(soc,"eutctId")
						eutctid.text=str(spor_id[organ_system])
						version=etree.SubElement(soc,"version")
						version.text=str(spor_term_version[organ_system])
						assessmth=etree.SubElement(nsae,"assessmentMethod")
						dict_override=etree.SubElement(nsae,"dictionaryOverridden")
						dict_override.text='false'
						value=etree.SubElement(assessmth,"value")
						value.text="ADV_EVT_ASSESS_TYPE.systematic"
						values=etree.SubElement(nsae,"values")

				value=etree.SubElement(values,"value")
				rpname=row["Treatment_Groups"]
				value.set("reportingGroupId","ReportingGroup-"+reporting_groups[row["Treatment_Groups"]])
				occurrences=etree.SubElement(value,"occurrences")
				occurrences.text=str(int(row["Occurrences_All_Number"]))
				subs_affected=etree.SubElement(value,"subjectsAffected")
				subs_affected.text=str(row["Subjects_Affected_Number"])
				subs_exposed=etree.SubElement(value,"subjectsExposed")
				subs_exposed.text=str(int(row["Subjects_Exposed_Number"]))

###################################################################
# Build serious adverse events XML
#
# node = parent node with tag <ae:adverseEvents>
# df_terms = dataframe loaded from serious ae spreadsheet
# reporting_groups = dictionary of reportting groups with ID  #
###################################################################      
def build_xml_serious_events(node, df_terms, reporting_groups):
    
		# serious adverse events
		saes=etree.SubElement(node,"seriousAdverseEvents")
		organ_system=""
		pref_term=""
		for index, row in df_terms.iterrows():
				if organ_system != row["System_Organ_Class"] or pref_term != row["Preferred_Term"]:
						organ_system = row["System_Organ_Class"]
						pref_term = row["Preferred_Term"]
						sae=etree.SubElement(saes,"seriousAdverseEvent")    
						term=etree.SubElement(sae,"term")
						term.text=pref_term
						soc=etree.SubElement(sae,"organSystem")
						eutctid=etree.SubElement(soc,"eutctId")
						eutctid.text=str(spor_id[organ_system])
						version=etree.SubElement(soc,"version")
						version.text=str(spor_term_version[organ_system])
						assessmth=etree.SubElement(sae,"assessmentMethod")
						dict_override=etree.SubElement(sae,"dictionaryOverridden")
						dict_override.text='false'
						value=etree.SubElement(assessmth,"value")
						value.text="ADV_EVT_ASSESS_TYPE.systematic"
						values=etree.SubElement(sae,"values")

				value=etree.SubElement(values,"value")
				rpname=row["Treatment_Groups"]
				#print (reporting_groups[rpname])
				value.set("reportingGroupId","ReportingGroup-"+reporting_groups[row["Treatment_Groups"]])
				occurrences=etree.SubElement(value,"occurrences")
				occurrences.text=str(int(row["Occurrences_All_Number"]))
				subs_affected=etree.SubElement(value,"subjectsAffected")
				subs_affected.text=str(int(row["Subjects_Affected_Number"]))
				subs_exposed=etree.SubElement(value,"subjectsExposed")
				subs_exposed.text=str(int(row["Subjects_Exposed_Number"]))

				ocrttn=etree.SubElement(value,"occurrencesCausallyRelatedToTreatment")
				if not pd.isna(row["Occurrences_Causally_Related_to_Treatment_Number"]):
				    ocrttn.text=str(int(row["Occurrences_Causally_Related_to_Treatment_Number"]))
				else:
				    ocrttn.text="0"
				fatalities=etree.SubElement(value,"fatalities")
				deaths=etree.SubElement(fatalities,"deaths")
				if not pd.isna(row["Fatalities_Number"]):
				    deaths.text=str(int(row["Fatalities_Number"]))
				else:   
				    deaths.text="0"
				dcrtt=etree.SubElement(fatalities,"deathsCausallyRelatedToTreatment")            
				if not pd.isna(row["Fatalities_Causally_Related_to_Treatment_Number"]):
				    dcrtt.text=str(int(row["Fatalities_Causally_Related_to_Treatment_Number"]))
				else:
				    dcrtt.text="0"

###################################################################
# Main function
#
# Read Spreadsheets into dataframe
# Create dataframe with overall counts for reporting groups section
# Create dataframe with serious adverse events
# Create dataframe wtih non-serious adverse events
###################################################################      
def build_XML():
	
		# load non-serious adverse event data from xlsx file to pandas dataframe
		df_nsae = pd.read_excel(directory + nsae_xlsx, sheet_name=0)
		df_nsae.Treatment_Groups=df_nsae.Treatment_Groups.str.replace('~   ', '')

		# Get rows with overall count by reporting group.   Overall count rows have System Organ Class = null
		df_nsae_counts=df_nsae[df_nsae.System_Organ_Class.isnull()]
		df_nsae_counts=df_nsae_counts[["Treatment_Groups","Subjects_Affected_Number","Subjects_Exposed_Number"]]
		df_nsae_counts.rename(columns={"Subjects_Affected_Number": "NSAE_Subjects_Affected"}, inplace=True)
		
		# Get rows with individual terms into dataframe
		df_nsae_terms=df_nsae[df_nsae.System_Organ_Class.notnull()]
		df_nsae_terms.Occurrences_All_Number = df_nsae_terms.Occurrences_All_Number.astype('int64')
		df_nsae_terms.MedDRA_Version = df_nsae_terms.MedDRA_Version.astype('int64')

		# load serious adverse event data from xlsx file to pandas dataframe
		df_sae = pd.read_excel(directory + sae_xlsx, sheet_name=0)
		df_sae.Treatment_Groups=df_sae.Treatment_Groups.str.replace('~   ', '')		

		# Get rows with overall count by reporting group
		df_sae_counts=df_sae[df_sae.System_Organ_Class.isnull()]
		df_sae_counts=df_sae_counts[["Treatment_Groups","Subjects_Affected_Number","Fatalities_Number","Fatalities_Causally_Related_to_Treatment_Number"]]
		df_sae_counts.rename(columns={"Subjects_Affected_Number": "SAE_Subjects_Affected"}, inplace=True)
		df_sae_counts.Treatment_Groups=df_sae_counts.Treatment_Groups.str.replace('~   ', '')
		print (df_sae_counts)
		
		# Get rows with individual terms into dataframe			
		df_sae_terms=df_sae[df_sae.System_Organ_Class.notnull()]
		df_sae_terms.MedDRA_Version = df_sae_terms.MedDRA_Version.astype('int64')
		
		# Merge SAE and NSAE count dataframes to get all counts we need into one row on a dataframe
		df_combined_counts=pd.merge(how='inner', left=df_sae_counts, right=df_nsae_counts, left_on='Treatment_Groups', right_on='Treatment_Groups')
		#df_combined_counts.Fatalities_Number = df_combined_counts.Fatalities_Number.astype('int64')

		# create adverse event element which is root element of entire file
		ae=etree.Element("{%s}adverseEvents" % (ns_eudra_ae), nsmap = NSMAP)
		thrshold=etree.SubElement(ae,"nonSeriousEventFrequencyThreshold")
		
		# non-serious event threshhold is always 5
		thrshold.text="5"
		timeframe=etree.SubElement(ae,"timeFrame")
		timeframe.text="timeframe text"
	
		# assessment method is always the same, thus hardcoded
		assessmth=etree.SubElement(ae,"assessmentMethod")
		value=etree.SubElement(assessmth,"value")
		value.text="ADV_EVT_ASSESS_TYPE.systematic"

		# dictionary info 
		dictionary=etree.SubElement(ae,"dictionary")
		othername=etree.SubElement(dictionary,"otherName")
		othername.set("{%s}nil" % (ns_xsi), "true")
		version=etree.SubElement(dictionary,"version")
		version.text=str(spor_version)
		dictname=etree.SubElement(dictionary,"name")
		value=etree.SubElement(dictname,"value")
		value.text="ADV_EVT_DICTIONARY_NAME.meddra"

		# create dictionary of treatment group names to reporting group id.  used later to tie reporting groups in individual terms
		reporting_groups = {}
		standard_cols=["adverseEventType","assessmentType","additionalDescription","organSystemName","sourceVocabulary","term"]
		for index, row in df_combined_counts.iterrows():
		    reporting_groups[row["Treatment_Groups"]]=str(index)		
		   
		# build XML nodes
		build_xml_reporting_groups(ae, df_combined_counts)
		build_xml_non_serious_events(ae, df_nsae_terms, reporting_groups)
		build_xml_serious_events(ae, df_sae_terms, reporting_groups)
		
		# write XML out to file
		write_xml(ae)
		
build_XML()

