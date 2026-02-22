import os
import io
import psycopg2
import openpyxl
import warnings
import shutil
import zipfile
import subprocess
from google import genai
from datetime import datetime
from google.genai import types


def extract_identity_documents(documents: list[dict]):
    """
    Ekstrak semua detail data yang ditemukan di setiap halaman pada file pdf dan berikan output seperti format dibawah.

    Struktur result seperti ini.
    result = [
        {"doc_type": doc_type, "data": sesuaikan dengan format args},
        {"doc_type": doc_type, "data": sesuaikan dengan format args}
    ]
    Isilah value doc type dan data sesuai args dibawah.
    Args:
        1. doc_type = formulir_pre_screening
           data = {
                "nama": "nama debitur",
                "alamat_rumah": "alamat rumah",
                "alamat_usaha": "alamat usaha",
                "bidang_usaha": "bidang usaha",
                "jumlah_permohonan_kredit": "nilai kredit yang diminta debitur",
                "tujuan_penggunaan_kredit": "tujuan penggunaan kredit"
           }
        2. doc_type = ktp_debitur
           data = {
                "nama": "nama pemilik ktp",
                "nomor_ktp": "nomor ktp",
                "tanggal_lahir": "tanggal lahir",
                "alamat": "alamat",
                "ktp_status": "ACCEPTED or REJECTED berdasarkan kesamaan informasi terhadap formulir_pre_screening"
           }
        3. doc_type = kartu_keluarga
           data = {
                "nomor_kartu_keluarga": "nomor dokumen",
                "kartu_keluarga_status": "ACCEPTED or REJECTED berdasarkan kesamaan informasi terhadap formulir_pre_screening"
           }
        4. doc_type = bpjs_kesehatan
           data = {
                "nama": "nama pemilik pada dokumen",
                "nomor_bpjs": "nomor dokumen",
                "tanggal_lahir": "tanggal lahir",
                "bpjs_kesehatan_status": "ACCEPTED or REJECTED berdasarkan kesamaan informasi terhadap formulir_pre_screening"
           }
    """
    pass


def postgresql_connect():
    """
    Connect to postgresql database
    """
    database_client = psycopg2.connect(
        database=os.getenv("DB_NAME"),
        user=os.getenv("DB_USERNAME"),
        password=os.getenv("DB_PASSWORD"),
        host=os.getenv("DB_HOST"),
        port=os.getenv("DB_PORT")
    )

    return database_client


def zip_file(user_temporary_dir):
    list_files = [file for file in os.listdir(user_temporary_dir) if ".xlsx" in file or ".pdf" in file]

    zip_file_path = f"{user_temporary_dir}/Extraction_Result.zip"

    with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename in list_files:
            zip_file.write(f"{user_temporary_dir}/{filename}", arcname=filename)

    # Delete existing xlsx and pdf files
    for filename in list_files:
        os.remove(f"{user_temporary_dir}/{filename}")

    # Move zip to memory
    with open(zip_file_path, "rb") as openfile:
        zip_memory = io.BytesIO(openfile.read())

    print("Successfully zipping file !")

    return zip_memory


def write_xlsx_and_pdf(update_data,
                       user_temporary_dir,
                       additional_data):
    warnings.filterwarnings("ignore")
    # xlsx processing
    worksheet = openpyxl.load_workbook(os.getenv("TEMPLATE_DOCUMENT_FILE_PATH"))
    sheet = worksheet.active
    for data in update_data:
        cell_code = data["cell_code"]
        cell_start = int(data["cell_start_idx"])
        current_cell_idx = cell_start
        for val in data["data"]:
            current_cell_name = f"{cell_code}{current_cell_idx}"
            sheet[current_cell_name] = val

            current_cell_idx += 1
    sheet["C5"] = str(additional_data["name"])
    sheet["C6"] = str(additional_data["position"])
    sheet["C7"] = str(additional_data["unit"])

    sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = False
    sheet.print_area = None

    worksheet.save(f"{user_temporary_dir}/extraction_result.xlsx")
    print("Successfully writing data to excel file !")

    # pdf processing
    SOFFICE_PATH = "/opt/homebrew/bin/soffice"
    command = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        f"{user_temporary_dir}/extraction_result.xlsx",
        "--outdir", user_temporary_dir
    ]

    subprocess.run(command, check=True)
    print("Successfully writing data to pdf file !")
    


def update_data_to_sheet(update_data,
                         additional_data):
    """
    Write the extracted data to worksheet loaded by openpyxl (to keep the format the same like the
    original format)
    """
    # Create user temporary dir
    unique_identifier = datetime.now().strftime("%f")
    dir_name = f"extraction_result_{unique_identifier}"
    os.mkdir(dir_name)

    write_xlsx_and_pdf(
        update_data,
        dir_name,
        additional_data
    )

    zip_memory = zip_file(dir_name)

    # Delete current working dir
    shutil.rmtree(dir_name)

    return zip_memory


def debitur_information(response_llm):
    """
    This function is used to construct the data for the debitur information (section I on spreadsheet)
    within the spreadsheet "checklist KUR"
    """
    debitur_name = "-"
    business_group = "-"
    debitur_address = "-"
    debitur_company_name = "-"
    debitur_company_address = "-"
    company_owner_name = "-"
    company_owner_address = "-"
    business_type = "-"
    amount_debt = "-"
    total_facilities = "-"
    debt_purpose = "-"
    for element_data in response_llm["documents"]:
        if "pre_screening" in element_data["doc_type"].lower():
            debitur_name = element_data["data"]["nama"] if "nama" in element_data["data"].keys() else "-"
            debitur_address = element_data["data"]["alamat_rumah"] if "alamat_rumah" in element_data["data"].keys() else "-"
            debitur_company_address = element_data["data"]["alamat_usaha"] if "alamat_usaha" in element_data["data"].keys() else"-"
            business_type = element_data["data"]["bidang_usaha"] if "bidang_usaha" in element_data["data"].keys() else "-"
            amount_debt = element_data["data"]["jumlah_permohonan_kredit"] if "jumlah_permohonan_kredit" in element_data["data"].keys() else "-"
            debt_purpose = element_data["data"]["tujuan_penggunaan_kredit"] if "tujuan_penggunaan_kredit" in element_data["data"].keys() else "-"
  

    # construct list values
    list_values = [
        {
            "cell_code": "C",
            "cell_start_idx": 16,
            "data": [
                debitur_name,
                business_group,
                debitur_address,
                debitur_company_name,
                debitur_company_address,
                company_owner_name,
                company_owner_address,
                business_type,
                amount_debt,
                total_facilities,
                debt_purpose
            ]
        }
    ]


    return list_values


def administration_information(response_llm):
    """
    This function is used to construct the data for the administration requirements (section II on spreadsheet)
    """
    ktp_debitur_information = "-"
    ktp_administrator_information = "-"
    ktp_collateral_owner_information = "-"
    ktp_debitur_status = "-"
    family_card_information = "-"
    family_card_status = "-"
    slik_jk = "-"
    black_list_national = "-"
    independent_appraisal = "-"
    public_accountant = "-"
    others = "-"
    for element_data in response_llm["documents"]:
        if "ktp_debitur" in element_data["doc_type"].lower():
            ktp_name_debitur = element_data["data"]["nama"] if "nama" in element_data["data"].keys() else "-"
            ktp_number_debitur = element_data["data"]["nomor_ktp"] if "nomor_ktp" in element_data["data"].keys() else "-"
            ktp_debitur_status = "Comply" if "ktp_status" in element_data["data"].keys() and element_data["data"]["ktp_status"] == "ACCEPTED" else "Not Comply"
            if ktp_name_debitur != "-" and ktp_number_debitur != "-":
                ktp_debitur_information = f"KTP debitur dengan nomor {ktp_number_debitur} atas nama {ktp_name_debitur} masih berlaku"
        elif "kartu_keluarga" in element_data["doc_type"].lower():
            family_card_number = element_data["data"]["nomor_kartu_keluarga"] if "nomor_kartu_keluarga" in element_data["data"].keys() else "-"
            family_card_status = "Comply" if "kartu_keluarga_status" in element_data["data"].keys() and element_data["data"]["kartu_keluarga_status"] == "ACCEPTED" else "Not Comply"
            if family_card_number != "-":
                family_card_information = f"KK pemilik agunan terlampir dengan nomor {family_card_number}"

    # Construct final list values
    list_values = [
        {
            "cell_code": "C",
            "cell_start_idx": 30,
            "data": [
                ktp_debitur_status,
                "-",
                "-",
                family_card_status,
                "-",
                "-",
                "-",
                "-",
                "-"
            ]
        },
        {
            "cell_code": "D",
            "cell_start_idx": 30,
            "data": [
                ktp_debitur_information,
                ktp_administrator_information,
                ktp_collateral_owner_information,
                family_card_information,
                slik_jk,
                black_list_national,
                independent_appraisal,
                public_accountant,
                others
            ]
        }
    ]

    return list_values



def extraction(file_bytes,
               additional_data):
    print("Start extracting data")
    # Define the google genai client
    genai_client = genai.Client(api_key=os.getenv("GOOGLE_CLOUD_API_KEY"))

    # Define tools and config
    tools = [extract_identity_documents]
    config = types.GenerateContentConfig(tools=tools)
    
    response = genai_client.models.generate_content(
        model="gemini-2.5-flash",
        contents=[
            types.Part.from_bytes(
                data=file_bytes,
                mime_type="application/pdf"
            )
        ],
        config=config
    )

    response_result = response.automatic_function_calling_history[1].parts[0].function_call.args

    update_data = []
    # Get debitur and administration information
    list_values_debitur_information = debitur_information(response_result)
    list_values_adm_information = administration_information(response_result)

    # Store the data on the list
    update_data += list_values_debitur_information + list_values_adm_information

    # Update data to spreadsheet
    zip_memory = update_data_to_sheet(update_data, additional_data)


    return zip_memory