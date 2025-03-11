from slims.slims import Slims
from slims.criteria import equals
from slims.criteria import conjunction, equals, contains, starts_with

slims = Slims("Content", "https://osdr-slims.smce.nasa.gov/slimsrest/", "{UN}", "{PW}", verify=False)




def create_rna_record_from_tissue(parent_tissue_sample_name, new_extract_type, new_extract_sample_name, new_extract_cntn_id, sample_storage_temp_label):
    """
    Creates an RNA record derived from an existing tissue sample.

    """
    try:
        # Fetch reference SLIMS objects
        print("fetching sample")
        tissue_samples = slims.fetch("Content", equals("cntn_cf_sampleName", parent_tissue_sample_name))
        print("fetching extract")
        extract_type = slims.fetch("ContentType", equals("cntp_name", new_extract_type))
        print("fetching temp")
        weight_test = slims.fetch("Test", equals("test_label", "weight"))
        #temp_type = slims.fetch("Content", contains("cntn_cf_sampleStoragetemp", sample_storage_temp_label))
        #print("temp type: ",temp_type)
        statuses = slims.fetch("Status", conjunction()
                       .add(equals("stts_name", "Pending"))
                       .add(equals("stts_fk_table", 2)))

        if not tissue_samples:
            print("No tissue sample found with name:", parent_tissue_sample_name)
            return

        tissue_sample = tissue_samples[0]
        print(f"Found tissue sample: {tissue_sample.cntn_cf_sampleName.value}")

        # Fetch the sample storage temp object



        # Define new record's field values
        rna_record_values = {
            'cntn_cf_sampleName': new_extract_sample_name,
            'cntn_fk_contentType': extract_type[0].pk(),
            'cntn_fk_location': tissue_sample.cntn_fk_location.value,
            'cntn_cf_fk_mission': tissue_sample.cntn_cf_fk_mission.value,
            'cntn_fk_status': statuses[0].pk(),
            'cntn_fk_originalContent': tissue_sample.pk(),  # Parent primary key
            'cntn_id': new_extract_cntn_id  # Should this be parent id? yes
        }

        # Create record in SLIMS
        created_rna = slims.add("Content", rna_record_values)

        created_result = slims.add("Result",
                           {'rslt_fk_content': created_rna.pk(),
                            'rslt_fk_test': weight_test[0].pk(),
                            'rslt_value': {
                                'amount': 0.02,
                                'unit_display': weight_test[0].follow("test_fk_resultTablefield")
                                                              .follow("tbfl_fk_unit")
                                                              .unit_abbreviation.value,
                                'unit_pk': weight_test[0].follow("test_fk_resultTablefield")
                                                         .tbfl_fk_unit.value}})
        
        print(f"New {new_extract_type} record '{new_extract_sample_name}' created successfully and linked to tissue sample '{parent_tissue_sample_name}'.")

    except Exception as e:
        print(f"Error creating RNA record: {str(e)}")
