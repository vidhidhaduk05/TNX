"""
Curated clean labels + code-system map for the cSDH TriNetX tables.

Labels are transcribed VERBATIM from the TriNetX PDF characteristic tables (visual
reads of pages 12/15/16/19 for Sx and the equivalent pages for MM). Codes are the
ground-truth key from the coordinate extractor. Code-system labels follow the UMLS
prefix literally (R5): UMLS:ICD10CM: -> ICD-10; UMLS:ICD10PCS: -> ICD-10-PCS;
TNX: -> TNX (TriNetX-curated lab); RxNorm numeric drug codes -> RxNorm.

The characteristics table shows medications and the platelet lab as BARE numeric codes
(RxNorm for drugs, TNX:9020 for platelets). We keep them faithful.
"""

# code -> clean display label (as printed in the PDF characteristic table)
LABELS = {
    # Demographics
    "AI": "Age at Index",
    "F": "Female",
    # Diagnosis / Comorbidities (ICD-10-CM)
    "I10": "Essential (primary) hypertension",
    "E08-E13": "Diabetes mellitus",
    "Z87.891": "Personal history of nicotine dependence",
    "I48": "Atrial fibrillation and flutter",
    "F17.210": "Nicotine dependence, cigarettes, uncomplicated",
    "I20-I25": "Ischemic heart diseases",
    "R51": "Headache",
    "R26": "Abnormalities of gait and mobility",
    "C00-D49": "Neoplasms",
    "I62.03": "Nontraumatic chronic subdural hemorrhage",
    "I62.00": "Nontraumatic subdural hemorrhage, unspecified",
    "I62.02": "Nontraumatic subacute subdural hemorrhage",
    "G81": "Hemiplegia and hemiparesis",
    "G40": "Epilepsy and recurrent seizures",
    "E78": "Disorders of lipoprotein metabolism and other lipidemias",
    "N18": "Chronic kidney disease (CKD)",
    "F10": "Alcohol related disorders",
    "K70-K77": "Diseases of liver",
    "G93.4": "Other and unspecified encephalopathy",
    "G93.5": "Compression of brain",
    "G30": "Alzheimer's disease",
    "G31.0": "Frontotemporal dementia",
    "F01": "Vascular dementia",
    "F02": "Dementia in other diseases classified elsewhere",
    "F03": "Unspecified dementia",
    "D65": "Disseminated intravascular coagulation [defibrination syndrome]",
    "D66": "Hereditary factor VIII deficiency",
    "D67": "Hereditary factor IX deficiency",
    "D68.0": "Von Willebrand disease",
    "D68.1": "Hereditary factor XI deficiency",
    "D68.2": "Hereditary deficiency of other clotting factors",
    "D68.3": "Hemorrhagic disorder due to circulating anticoagulants",
    "D68.4": "Acquired coagulation factor deficiency",
    "D68.5": "Primary thrombophilia",
    "D68.6": "Other thrombophilia",
    "D68.8": "Other specified coagulation defects",
    "D68.9": "Coagulation defect, unspecified",
    "D69.0": "Allergic purpura",
    "D69.1": "Qualitative platelet defects",
    "D69.2": "Other nonthrombocytopenic purpura",
    "D69.3": "Immune thrombocytopenic purpura",
    "D69.4": "Other primary thrombocytopenia",
    "D69.5": "Secondary thrombocytopenia",
    "D69.6": "Thrombocytopenia, unspecified",
    "D69.8": "Other specified hemorrhagic conditions",
    "D69.9": "Hemorrhagic condition, unspecified",
    # Medication (RxNorm) -> generic drug name
    "1364430": "apixaban",
    "11289": "warfarin",
    "1114195": "rivaroxaban",
    "1191": "aspirin",
    "32968": "clopidogrel",
    "613391": "prasugrel",
    "1116632": "ticagrelor",
    "3521": "dipyridamole",
    # Laboratory
    "9020": "Platelets [#/volume] in Blood",
}

# code-system label per code (from UMLS prefix, R5)
def code_system(code):
    """Return the display system tag for a code, per its UMLS prefix / type."""
    if code in ("AI",):
        return "TNX"          # Age at Index is a TriNetX demographic token (shown as 'A1')
    if code == "F":
        return None            # Female = demographic, no coding system
    if code == "9020":
        return "TNX"          # TNX:9020 platelet lab
    # ICD-10-CM diagnosis codes: letter+digits, ranges, or dotted
    import re
    if re.match(r"^[A-Z]\d", code):
        return "ICD-10"
    if code.isdigit():
        return "RxNorm"        # bare numeric medication codes
    return None


# Platelet range-bin labels (Laboratory section, categorical). Key by the bin text
# the extractor reconstructs; render open/closed faithfully.
# PLT<50 studies use the first 5 bins (0-10 ... 40-50); PLT<100 studies use all 10
# (0-10 ... 90-100). The extractor keys on the actual bin text it reconstructs, so
# this list is a reference/sanity set only; pretty_bin() handles any numeric range.
RANGE_BINS = ["0 - 10 10*3/uL", "10 - 20 10*3/uL", "20 - 30 10*3/uL",
              "30 - 40 10*3/uL", "40 - 50 10*3/uL", "50 - 60 10*3/uL",
              "60 - 70 10*3/uL", "70 - 80 10*3/uL", "80 - 90 10*3/uL",
              "90 - 100 10*3/uL"]

# Pretty range-bin display (keep numeric faithful; tidy spacing/units).
# Use search (not match) and ignore any leading token such as a stray "Blood "
# that can bleed in from the preceding "Platelets [#/volume] in Blood" line on
# some PDFs (observed in the Sx+MMAE PLT<100 export).
def pretty_bin(raw):
    import re
    m = re.search(r"(\d+)\s*-\s*(\d+)\s*10\*3/uL", raw)
    if m:
        return f"{m.group(1)}\u2013{m.group(2)} \u00d710\u00b3/\u00b5L"
    return raw
