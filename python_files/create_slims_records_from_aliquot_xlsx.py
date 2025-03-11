from slims.slims import Slims
from slims.criteria import equals
from slims.criteria import conjunction, equals, contains, starts_with

slims = Slims("Content", "https://osdr-slims.smce.nasa.gov/slimsrest/", "{UN}", "{PW}", verify=False)



def create_rna_record_from_tissue(
    parent_tissue_sample_name,
    new_extract_type,
    new_extract_sample_name,
    new_extract_cntn_id):

    """
    Creates a new RNA record derived from an existing tissue sample.

    """
    try:
        # fetch reference SLIMS objects
        tissue_samples = slims.fetch("Content", equals("cntn_barCode", parent_tissue_sample_name))
        extract_type = slims.fetch("ContentType", equals("cntp_name", new_extract_type))
        statuses = slims.fetch("Status", conjunction()
                       .add(equals("stts_name", "Pending"))
                       .add(equals("stts_fk_table", 2)))

        if not tissue_samples:
            print("No tissue sample found with name:", parent_tissue_sample_name)
            return

        tissue_sample = tissue_samples[0]
        print(f"Found tissue sample: {tissue_sample.cntn_barCode.value}")

        # define new record's field values
        rna_record_values = {
            'cntn_cf_sampleName': new_extract_sample_name,
            'cntn_fk_contentType': extract_type[0].pk(),
#            'cntn_fk_location': tissue_sample.cntn_fk_location.value,
            'cntn_cf_fk_mission': tissue_sample.cntn_cf_fk_mission.value,
            'cntn_fk_status': statuses[0].pk(),
            'cntn_fk_originalContent': tissue_sample.pk(),  # Parent primary key
            'cntn_id': new_extract_cntn_id
        }

        created_rna = slims.add("Content", rna_record_values)
        print(f"New {new_extract_type} record '{new_extract_sample_name}' created successfully and linked to tissue sample '{parent_tissue_sample_name}'.")

    except Exception as e:
        print(f"Error creating RNA record: {str(e)}")

def get_value(row, column_name):
    value = row.get(column_name)
    return str(value) if pd.notna(value) and not value == "nan" else "NA"

def process_xlsx_and_upload_to_slims(file_path):
    df = pd.read_excel(file_path)

    for _, row in df.iterrows():
        # extract all required fields from the row
        parent_tissue_sample_name = get_value(row, 'Subject Barcode')
        new_extract_type = "RNA"  # assuming all extracts are RNA
        new_extract_sample_name = get_value(row, 'Extract ID')
        new_extract_cntn_id = get_value(row, 'GL Sample ID')

        # call the SLIMS function
        create_rna_record_from_tissue(
            parent_tissue_sample_name,
            new_extract_type,
            new_extract_sample_name,
            new_extract_cntn_id
        )

    print(f"\nAll records processed! {len(df)} samples created and uploaded to SLIMS.")

# example usage
file_path = 'output_extract_list_test_aq2.xlsx'  # replace with actual file path
process_xlsx_and_upload_to_slims(file_path)
