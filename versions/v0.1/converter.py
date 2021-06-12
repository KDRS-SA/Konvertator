import subprocess
import csv
import pathlib
import argparse
from datetime import datetime
import psutil
import pendulum
import extract_msg
import shlex

# Liste over pronomkoder som blir konvertert, denne kan nok erstattes eller flyttes ut av scriptet til en settings-fil.

pronom_type              = {}
pronom_type['fmt/39']    = {'Name': 'Microsoft Word Document 6.0/95', 'convert': 'libreoffice'}
pronom_type['fmt/136']   = {'Name': 'OpenDocument Text 1.0', 'convert': 'libreoffice'}
pronom_type['fmt/214']   = {'Name': 'Microsoft Excel for Windows', 'convert': 'libreoffice'}
pronom_type['fmt/126']   = {'Name': 'Microsoft Powerpoint Presentation', 'convert': 'libreoffice'}
pronom_type['fmt/215']   = {'Name': 'Microsoft Powerpoint for Windows', 'convert': 'libreoffice'}
pronom_type['fmt/487']   = {'Name': 'Macro Enabled Microsoft Powerpoint', 'convert': 'libreoffice'}
pronom_type['fmt/290']   = {'Name': 'OpenDocument Text ', 'convert': 'libreoffice'}
pronom_type['fmt/291']   = {'Name': 'OpenDocument Text ', 'convert': 'libreoffice'}
pronom_type['fmt/355']   = {'Name': 'Rich Text Format ', 'convert': 'libreoffice'}
pronom_type['fmt/40']    = {'Name': 'Microsoft Word Document ', 'convert': 'libreoffice'}
pronom_type['fmt/412']   = {'Name': 'Microsoft Word for Windows ', 'convert': 'libreoffice'}
pronom_type['fmt/445']   = {'Name': 'Microsoft Excel Macro-Enabled', 'convert': 'libreoffice'}
pronom_type['fmt/50']    = {'Name': 'Rich Text Format ', 'convert': 'libreoffice'}
pronom_type['fmt/523']   = {'Name': 'Macro enabled Microsoft Word Document OOXML ', 'convert': 'libreoffice'}
pronom_type['fmt/53']    = {'Name': 'Rich Text Format ', 'convert': 'libreoffice'}
pronom_type['fmt/61']    = {'Name': 'Microsoft Excel 97 Workbook (xls)', 'convert': 'libreoffice'}
pronom_type['fmt/597']   = {'Name': 'Microsoft Word Template ', 'convert': 'libreoffice'}
pronom_type['fmt/598']   = {'Name': 'Microsoft Excel Template ', 'convert': 'libreoffice'}
pronom_type['fmt/96']    = {'Name': 'Hypertext Markup Language', 'convert': 'libreoffice'}
pronom_type['fmt/99']    = {'Name': 'Hypertext Markup Language', 'convert': 'libreoffice'}
pronom_type['fmt/258']   = {'Name': 'Microsoft Works Word Processor 5-6', 'convert': 'libreoffice'}
pronom_type['fmt/609']   = {'Name': 'Microsoft Word (Generic) 6.0-2003', 'convert': 'libreoffice'}
pronom_type['fmt/38']    = {'Name': 'Microsoft Word for Windows Document 2.0', 'convert': 'libreoffice'}
pronom_type['fmt/595']   = {'Name': 'Microsoft Excel Non-XML Binary Workbook 2007 onwards', 'convert': 'libreoffice'}
pronom_type['fmt/59']    = {'Name': 'Microsoft Excel 5.0/95 Workbook (xls) 5/95', 'convert': 'libreoffice'}
pronom_type['x-fmt/88']  = {'Name': 'Microsoft Powerpoint Presentation 4.0', 'convert': 'libreoffice'}
pronom_type['fmt/295']   = {'Name': 'OpenDocument Spreadsheet 1.2', 'convert': 'libreoffice'}
pronom_type['fmt/294']   = {'Name': 'OpenDocument Spreadsheet 1.1', 'convert': 'libreoffice'}
pronom_type['fmt/57']    = {'Name': 'Microsoft Excel 4.0 Worksheet (xls)', 'convert': 'libreoffice'}

pronom_type['x-fmt/430'] = {'Name': 'Microsoft Outlook Email Message', 'convert': 'email'}

## Stats-logging og setter opp noen av standardinnstillingene.

results = {}
results['stats'] = {}
results['stats']['converted'] = 0
results['stats']['unconverted'] = 0

results_dir = 'results/'
pathlib.Path(results_dir).mkdir(parents=True, exist_ok=True)
log_file = open(f'{results_dir}logfile.txt','a')

## Funksjon for å sjekke filnavnskolisjoner som ikke har blitt tatt i bruk...

def filsjekk(filnavn, filtype):
    if pathlib.Path(f"{filnavn}.{filtype}").is_file() is True:
        fil_check = False
        counter = 0
        while fil_check is False:
            if pathlib.Path(f"{filnavn}_{counter}.{filtype}").is_file() is True:
                fil_check = True
                return f"{filnavn}_{counter}.{filtype}"
            else:
                counter += 1
    else:
        return f"{filnavn}.{filtype}"

def logging(tekst):
    global log_file
    date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#    print(date + '\t' + tekst)
    log_file.write(date + '\t' + tekst + '\n')

## Setter opp output-csv fra konverteringsprosessen, hvis filen allerede eksisterer så lager den et dictionary med de som allerede ligger i csven som brukes for å hoppe over filer. 
## Dette er gjort for å sikre sømløs restart av prosessen. Siden det er en prosess som kan ta litt tid.

already_converted = 0
proc_restart = False
if pathlib.Path(f'{results_dir}convert_output.csv').is_file() is True:
    proc_restart = True
    converted_files = {}
    with open(f'{results_dir}convert_output.csv', newline='\n') as convert_csv:
        convert_csv_reader = csv.reader(convert_csv, delimiter=',', quotechar='"')
        for row in convert_csv_reader:
            converted_files[row[0]] = row[1]
            already_converted += 1
    output_file = open(f'{results_dir}convert_output.csv', mode='a')
else:
    output_file = open(f'{results_dir}convert_output.csv', mode='w')
output_writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

## OK, denne funksjonen kan være litt overflødig, men... den teller opp antallet filer for å vite antallet totalt i konverteringsprosessen. 

with open(f'{results_dir}pronom_check.csv', newline='\n') as pronom_csv:
    pronom_csv_reader = csv.reader(pronom_csv, delimiter=',', quotechar='"')
    total_files = len(list(pronom_csv_reader))

## Dette er selve konverteringsfunksjonen
## Tredelt: 1. den konverterer med libreoffice de som kan konverteres (se tabellen i starten)
## 2. Den konverterer outlook-filer (.msg) til ren tekstfiler.
## 3. Den hopper over de som enten er i godkjent format eller i et format som ikke lar seg konvertere.

with open(f'{results_dir}pronom_check.csv', newline='\n') as pronom_csv:
    pronom_csv_reader = csv.reader(pronom_csv, delimiter=',', quotechar='"')
    for row in pronom_csv_reader:
        if proc_restart is True:
            if row[0] in converted_files:
                continue
        empty_file = 'n'
        if row[2] == '0':
            empty_file = 'y'
        filename = {}
        filename['old'] = pathlib.Path(row[0])
        filename['dir'] = filename['old'].parent
        if row[1] in pronom_type:
            if pronom_type[row[1]]['convert'] == 'libreoffice':
                try:
                    filename['new'] = pathlib.Path(filename['dir'].as_posix() + '/' + filename['old'].stem + '.pdf')
                    subprocess.run(['libreoffice --headless --convert-to pdf --outdir ' + shlex.quote(filename['dir'].as_posix()) + ' ' + shlex.quote(filename['old'].as_posix())], shell=True, stdout=subprocess.DEVNULL, timeout=360)
                    if filename['new'].is_file() is True:
                        filename['old'].unlink()
                    results['stats']['converted'] += 1
                    logging(f"{already_converted + results['stats']['converted'] + results['stats']['unconverted']}/{total_files}\t {filename['old']} converted")
                    output_writer.writerow([filename['old'], filename['new'], row[1], 'conv', empty_file])
                except subprocess.TimeoutExpired:
                    for proc in psutil.process_iter():
                        if proc.name() == 'soffice.bin':
                            proc.kill()
                    output_writer.writerow([filename['old'], filename['old'], row[1], 'time', empty_file])
                    results['stats']['unconverted'] += 1
                    logging(f"{already_converted +results['stats']['converted'] + results['stats']['unconverted']}/{total_files}\t {filename['old']} timed out")
            if pronom_type[row[1]]['convert'] == 'email':
                filename['new'] = pathlib.Path(filename['dir'].as_posix() + '/' + filename['old'].stem + '.txt')
                try:
                    msg = extract_msg.Message(filename['old'])
                    output_text = '==============================================================================\n'
                    output_text += f'Sendt:\t\t{msg.date} ({pendulum.parse(msg.date, strict=False).isoformat()})\n'
                    output_text += f'Avsender:\t{msg.sender}\n'
                    output_text += f'Mottaker(e):\t' + ', '.join([f'{x.name} <{x.email}>' for x in msg.recipients]) + '\n'
                    output_text += f'Vedlegg:\t' + ', '.join([f'<{x.longFilename}>' for x in msg.attachments]) + '\n'
                    output_text += f'Emne:\t{msg.subject}\n'
                    output_text += f'==============================================================================\n{msg.body}'
                    with open(filename['new'], 'w') as outlook_out:
                        outlook_out.write(output_text)
                    if filename['new'].is_file() is True:
                        filename['old'].unlink()
                    results['stats']['converted'] += 1
                    logging(f"{already_converted + results['stats']['converted'] + results['stats']['unconverted']}/{total_files}\t {filename['old']} converted")
                    output_writer.writerow([filename['old'], filename['new'], row[1], 'conv', empty_file])
                except:
                    results['stats']['unconverted'] += 1
                    output_writer.writerow([filename['old'], filename['old'], row[1], 'nconv', empty_file])
                    logging(f"{already_converted + results['stats']['converted'] + results['stats']['unconverted']}/{total_files}\t {filename['old']} unconverted")
        else:
            results['stats']['unconverted'] += 1
            output_writer.writerow([filename['old'], filename['old'], row[1], 'nconv', empty_file])
            logging(f"{already_converted + results['stats']['converted'] + results['stats']['unconverted']}/{total_files}\t {filename['old']} unconverted")
output_file.close()
logging('Done converting')
log_file.close()
