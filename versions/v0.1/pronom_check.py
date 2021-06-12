import subprocess
import csv
import pathlib
import argparse
import os
import sys
import json
from datetime import datetime

## Funksjon for å få ut scanne filene inne i mappen den får som input. Returnerer csv med filnavn, pronomkode, filstørrelse

def siegfriedtest(sieg_filename):
    global pronom_stats
    siegfried_output = []
    siegfried_list = subprocess.check_output(["/opt/siegfried/bin/sf", "-nr", "-csv", sieg_filename]).splitlines()
    for file in siegfried_list[1:]:
        siegfried_check = True
        try:
            file.decode('utf8')
        except UnicodeDecodeError:
            siegfried_check = False
        if siegfried_check is True:
            siegfriedobjekt_ny = csv.reader([file.decode('utf8')], delimiter=',', quotechar='"')
            for row in siegfriedobjekt_ny:
                file_output = []
                file_output.append(row[0])  # Filename
                file_output.append(row[5])  # Pronomkode
                file_output.append(row[1])  # Size
                siegfried_output.append(file_output)
                if row[5] not in pronom_stats['count']:
                    pronom_stats['names'][row[5]] = row[6]
                    if row[7] != '':
                        pronom_stats['names'][row[5]] += f' vers. {row[7]}'
                    pronom_stats['count'][row[5]] = 1
                else:
                    pronom_stats['count'][row[5]] += 1
        else:
            logging(f"UTF8 error - {ascii(file)}")
    return siegfried_output

def logging(tekst):
    global log_file
    date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_file.write(date + '\t' + tekst + '\n')

## Funksjon som gir prompt for å velge mellom systemforekomster for SIARD-filene, den finner filene som xmler i script_path/settings/*.xml

def xml_search(system):
    script_path = pathlib.Path(sys.argv[0]).parent.resolve()
    if system != '':
        if script_path.joinpath(script_path, f"settings/{system}.xml").exists() is True:
            return script_path.joinpath(script_path, f"settings/{system}.xml").as_posix()
        else:
            return xml_search('')
    else:
        count = 1
        path_dict = {}
        for path in pathlib.Path.joinpath(script_path, f"settings/").rglob('*'):
            path_dict[count] = {}
            path_dict[count]['path'] = path
            path_dict[count]['name'] = path.stem
            count += 1
        print('\nVelg systemforekomst:\n')
        for x in path_dict:
            print(f"\t[{x}]\t{path_dict[x]['name']}")
        print('')
        user_input = input('#: ')
        try:
            user_input_int = int(user_input)
        except ValueError:
            return xml_search('')
        if user_input_int in path_dict:
            return path_dict[user_input_int]['path'].as_posix()
        else:
            return xml_search('')

def prompt(prompt_input):
    print('\n')
    print(prompt_input)
    user_input = input(': ')
    print('\n')
    return user_input

## Command line arguments.

parser = argparse.ArgumentParser()
parser.add_argument("-d", "--docs", help="Dokumentmappen", action="store", default=False)
parser.add_argument("-i", "--siard", help="SIARD-filen", action="store", default=False)
parser.add_argument("-s", "--system", help="Hvilket system", action="store", default='')
args = parser.parse_args()

## Etablerer logging

results_dir = 'results/'
pathlib.Path(results_dir).mkdir(parents=True, exist_ok=True)
if pathlib.Path(f'{results_dir}logfile.txt').is_file() is True:
    pathlib.Path(f'{results_dir}logfile.txt').unlink()
log_file = open(f'{results_dir}logfile.txt','x')

pronom_stats = {}
pronom_stats['names'] = {}
pronom_stats['count'] = {}
filliste = []

## Prompts for userinput med systemforekomst, path til dokumentkatalogen og path til SIARD-filen. Settingene blir lagret i results/settings.json og viderebrukes av de andre scriptene.

if pathlib.Path(f'{results_dir}settings.json').is_file() is True:
    settings_dict = json.load(open(f"{results_dir}settings.json"))

else:
    settings_dict = {}
    settings_dict['system_file'] = xml_search(args.system)
    siard_file = args.siard
    if siard_file is not False:
        if pathlib.Path(siard_file).exists() is False:
            siard_file = False
    if siard_file is False:
        while siard_file is False:
            test_siard_file = prompt('Angi SIARD-fil:')
            if pathlib.Path(test_siard_file).exists() is True:
                siard_file = test_siard_file
    settings_dict['siard_file'] = pathlib.Path(siard_file).as_posix()
    document_path = args.docs
    if document_path is not False:
        if pathlib.Path(document_path).exists() is False:
            document_path = False
    if document_path is False:
        while document_path is False:
            test_document_path = prompt('Angi dokumentmappe:')
            if pathlib.Path(test_document_path).exists() is True:
                document_path = test_document_path
    settings_dict['document_path'] = pathlib.Path(document_path).as_posix()
    open(f'{results_dir}settings.json', 'w').write(json.dumps(settings_dict))

document_dir = pathlib.Path(settings_dict['document_path'])


## Rename-scan, et problem i Arkade har vært at den ødelegger utf-8-tegn. Denne scanneren scanner underfiler i dokument-path (som er satt over) og erstatter tegn til gyldige UTF8 og logger til loggfilen.

replace_table = {}
replace_table['\udce6'] = 'æ'
replace_table['\udcf8'] = 'ø'
replace_table['\udce5'] = 'å'
replace_table['\udcc6'] = 'Æ'
replace_table['\udcd8'] = 'Ø'
replace_table['\udcc5'] = 'Å'

for path in pathlib.Path('../').rglob('*'):
    try:
        path.as_posix().encode('utf8')
    except UnicodeEncodeError:
        file_string = path.as_posix()
        replace_chars = []
        for x in file_string:
            if x in replace_table:
                replace_chars.append(x)
        for y in replace_chars:
            file_string = file_string.replace(y, replace_table[y])
        logging(f"{file_string} corrected utf-8 chars")
        path.rename(file_string)

## Hovedpronomscanning. Den scanner gjennom dokument-path og setter i gang siegfriedsjekk-funksjonen. Resultatene lagres i en csv-fil for de etterfølgene scriptene.

logging(f"Started PRONOM-scan")
if pathlib.Path(f'{results_dir}pronom_check.csv').is_file() is True:
    pathlib.Path(f'{results_dir}pronom_check.csv').unlink()
scanned_number = 0
total_number = 0
for path in pathlib.Path(document_dir).rglob('*'):
    if path.is_dir() is False:
        total_number += 1
for path in pathlib.Path(document_dir).rglob('*'):
    if path.is_dir() is True:
        filliste.extend(siegfriedtest(path))
        scanned_number += len(filliste)
        print(f"{scanned_number} / {total_number} {ascii(path.as_posix())}")
    if filliste != []:
        with open(f'{results_dir}pronom_check.csv', mode='a') as output_file:
            writer = csv.writer(output_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            for row in filliste:
                writer.writerow(row)
        filliste = []

# Logge til results

if pathlib.Path(f'{results_dir}pronom_stats.json').is_file() is True:
    pathlib.Path(f'{results_dir}pronom_stats.json').unlink()
open(f'{results_dir}pronom_stats.json', 'w').write(json.dumps(pronom_stats))
logging(f"{scanned_number} files scanned for PRONOM")


