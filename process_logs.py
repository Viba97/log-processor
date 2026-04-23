import os
import glob
import sys
import pandas as pd

def main():
    # 1. Get the folder path from the command line argument or prompt the user
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = input("Please enter the path to the folder containing .Qnf files: ").strip()

    # Verify if the provided path exists
    if not os.path.exists(folder_path):
        print(f"Error: The directory '{folder_path}' does not exist.")
        sys.exit(1)

    # 2. Find all .Qnf files in that folder
    file_pattern = os.path.join(folder_path, '*.Qnf')
    file_paths = glob.glob(file_pattern)
    
    if not file_paths:
        print(f"No .Qnf files found in '{folder_path}'.")
        sys.exit(1)

    # List to hold the extracted data from all files
    all_files_data = []

    # 3. Iterate through each file and extract the information
    for file_path in file_paths:
        file_data = {}
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                # Skip empty lines or section headers like [Unit]
                if not line or line.startswith('['):
                    continue
                # Split the line at the first '=' sign
                if '=' in line:
                    key, value = line.split('=', 1)
                    file_data[key.strip()] = value.strip()
                    
        all_files_data.append(file_data)

    # 4. Convert the list of dictionaries into a pandas DataFrame
    df = pd.DataFrame(all_files_data)

    # 5. Separate 'TimeStamp' and insert before 'TimeElapsed'
    if 'TimeStamp' in df.columns and 'TimeElapsed' in df.columns:
        insert_loc = df.columns.get_loc('TimeElapsed')
        split_data = df['TimeStamp'].str.split(' ', n=1, expand=True)
        df.insert(insert_loc, 'date', split_data[0])
        df.insert(insert_loc + 1, 'time', split_data[1])
        df = df.drop('TimeStamp', axis=1)

    # 6. Save the dataframe to an Excel file in the parent directory
    parent_folder = os.path.dirname(folder_path)
    output_excel_path = os.path.join(parent_folder, 'extracted_data.xlsx')
    df.to_excel(output_excel_path, index=False, engine='openpyxl')
    
    # We will store our output text here to save it to a file later
    report_lines = []
    
    report_lines.append(f"Dati estratti e salvati in {output_excel_path}\n")

    # 7. Data elaboration for FPY and Global Yield
    df['DateTime'] = pd.to_datetime(df['date'] + ' ' + df['time'], format='%d.%m.%Y %H:%M:%S')
    df_sorted = df.sort_values(by=['SN', 'DateTime'])
    df_primi_test = df_sorted.drop_duplicates(subset=['SN'], keep='first')

    totale_schede_uniche = len(df_primi_test)
    pass_al_primo_colpo = len(df_primi_test[df_primi_test['Result'] == 'PASS'])
    fail_al_primo_colpo = len(df_primi_test[df_primi_test['Result'] == 'FAIL'])

    if totale_schede_uniche > 0:
        fpy_percentuale = (pass_al_primo_colpo / totale_schede_uniche) * 100
        ffy_percentuale = (fail_al_primo_colpo / totale_schede_uniche) * 100
    else:
        fpy_percentuale = 0.0
        ffy_percentuale = 0.0

    report_lines.append("--- RISULTATI: First Pass Yield ---")
    report_lines.append(f"Unità totali processate (senza duplicati): {totale_schede_uniche}")
    report_lines.append(f"Unità conformi al primo passaggio (PASS): {pass_al_primo_colpo}")
    report_lines.append(f"Unità scartate al primo passaggio (FAIL): {fail_al_primo_colpo}")
    report_lines.append(f"First Pass Yield (FPY): {pass_al_primo_colpo}/{totale_schede_uniche} = {fpy_percentuale:.2f}%")
    report_lines.append(f"First FAIL Yield: {fail_al_primo_colpo}/{totale_schede_uniche} = {ffy_percentuale:.2f}%\n")

    # Serial Number testati più di una volta
    conteggi_seriali = df['SN'].value_counts()
    seriali_ripetuti = conteggi_seriali[conteggi_seriali > 1]

    report_lines.append("--- SCHEDE TESTATE PIÙ DI UNA VOLTA ---")
    if not seriali_ripetuti.empty:
        report_lines.append(f"Sono state trovate {len(seriali_ripetuti)} schede con test multipli:")
        for sn, conteggio in seriali_ripetuti.items():
            report_lines.append(f"- Serial Number: {sn} (Testato {conteggio} volte)")
    else:
        report_lines.append("Nessun Serial Number è stato testato più di una volta.")
    report_lines.append("\n")

    # FAIL globale su tutti i record 
    totale_records = len(df)
    totale_fail_globali = len(df[df['Result'] == 'FAIL'])

    if totale_records > 0:
        fail_globale_percentuale = (totale_fail_globali / totale_records) * 100
    else:
        fail_globale_percentuale = 0.0

    report_lines.append("--- RISULTATI: Percentuale FAIL Globale (Punto G) ---")
    report_lines.append(f"Record totali registrati (inclusi duplicati): {totale_records}")
    report_lines.append(f"Record totali con esito FAIL: {totale_fail_globali}")
    report_lines.append(f"FAIL Globale: {totale_fail_globali}/{totale_records} = {fail_globale_percentuale:.2f}%")

    # 8. Print everything to the console AND save it to a text file
    final_report = '\n'.join(report_lines)
    print(final_report)

    output_txt_path = os.path.join(parent_folder, 'report_risultati.txt')
    with open(output_txt_path, 'w', encoding='utf-8') as report_file:
        report_file.write(final_report)
    
    print(f"\n[INFO] Il report testuale è stato salvato in: {output_txt_path}")

if __name__ == "__main__":
    main()