import csv
import pathlib
import argparse
import zipfile
import shutil
import json
from datetime import datetime
from lxml import etree

#parser = argparse.ArgumentParser()
#parser.add_argument("-i", "--input", help="Siardfil", action="store")
#parser.add_argument("-d", "--docs", help="Dokumentmappen", action="store")
#args = parser.parse_args()



namespace = {'siard': 'http://www.bar.admin.ch/xmlns/siard/2/table.xsd',
             'siard-metadata': 'http://www.bar.admin.ch/xmlns/siard/2/metadata.xsd',
             'xml-schema': 'http://www.w3.org/2001/XMLSchema'
            }

results_dir = 'results/'
temp_dir = 'temp/'
pathlib.Path(temp_dir).mkdir(parents=True, exist_ok=True)
settings_dict = json.load(open(f"{results_dir}settings.json"))

document_dir = pathlib.Path(settings_dict['document_path'])
siard_filename = pathlib.Path(settings_dict['siard_file'])

# Import xml-fil for systemvariant
system_xml = etree.parse(settings_dict['system_file'])
system_metadata = {}
system_metadata['tablename'] = system_xml.xpath('/xml/metadata/tablename')[0].text
system_metadata['sqltype'] = system_xml.xpath('/xml/metadata/sql/type')[0].text
system_metadata['sqltypeOriginal'] = system_xml.xpath('/xml/metadata/sql/typeOriginal')[0].text
table_name = system_metadata['tablename']
log_file = open(f'{results_dir}logfile.txt','a')

def add_node(filename, nodename):
    node = etree.Element('{' + str(namespace['siard']) + '}' + nodename)
    node.text = filename
    return(node)

def add_node_metadataxml(nodename):
    node = etree.Element('{' + str(namespace['siard-metadata']) + '}' + 'column')
    etree.SubElement(node, '{' + str(namespace['siard-metadata']) + '}name').text = nodename
    etree.SubElement(node, '{' + str(namespace['siard-metadata']) + '}type').text = system_metadata['sqltype']
    etree.SubElement(node, '{' + str(namespace['siard-metadata']) + '}typeOriginal').text = system_metadata['sqltypeOriginal']
    etree.SubElement(node, '{' + str(namespace['siard-metadata']) + '}nullable').text = 'true'
    return(node)

def add_node_xsd(nodename):
    node = etree.Element('{' + str(namespace['xml-schema']) + '}element')
    node.attrib['name'] = nodename
    node.attrib['minOccurs'] = '0'
    node.attrib['type'] = 'xs:string'
    return(node)

def logging(tekst):
    global log_file
    date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(date + '\t' + tekst)
    log_file.write(date + '\t' + tekst + '\n')

# Åpner Siardfil
zip = zipfile.ZipFile(siard_filename)
logging('Unpacking SIARD-file')
zip.extractall(temp_dir)

# Henter riktig tabell direkte fra basen
root_metadata = etree.parse(temp_dir + "header/metadata.xml")
doktable = root_metadata.xpath('//siard-metadata:name[text()="' + table_name +'"]/..', namespaces=namespace)[0]
table = doktable.xpath('siard-metadata:folder', namespaces=namespace)[0].text
metadata_col = 'c' + str(len(doktable.xpath('siard-metadata:columns/siard-metadata:column', namespaces=namespace)) + 1)

# Leser ut tabellen
logging('Processing SIARD, this might take a while...')
siard_table = f'content/schema0/{table}/{table}.xml'
root = etree.parse(temp_dir + siard_table)

# Laster inn dictionary med filene
converted_files = {}
with open(f'{results_dir}convert_output.csv', newline='\n') as convert_csv:
    convert_csv_reader = csv.reader(convert_csv, delimiter=',', quotechar='"')
    for row in convert_csv_reader:
        converted_files[row[0]] = {}
        converted_files[row[0]]['old'] = row[0]
        converted_files[row[0]]['new'] = row[1]
        converted_files[row[0]]['pronom'] = row[2]
        converted_files[row[0]]['conv'] = row[3]
        converted_files[row[0]]['siard'] = 'n'
        converted_files[row[0]]['empty'] = row[4]

failed_ref = []

siard_refs = 0
#Konstruerer opp filnavn fra SIARD-filen.
for x in root.xpath('siard:row', namespaces=namespace):
    siard_refs += 1
    file_name = ''
    for node in system_xml.xpath('/xml/path/node'):
        if 'col' in node.attrib:
            if x.xpath(f"siard:{node.attrib['col']}", namespaces=namespace) != [] and x.xpath(f"siard:{node.attrib['col']}", namespaces=namespace)[0].text is not None:
                file_name += x.xpath(f"siard:{node.attrib['col']}", namespaces=namespace)[0].text
        if 'text' in node.attrib:
            file_name += node.attrib['text']
    file_name = pathlib.Path(file_name.replace('\\u005c', '/'))
    file_path = pathlib.Path(document_dir / file_name)
    if file_path.as_posix() in converted_files:
        converted_files[file_path.as_posix()]['siard'] = 'y'
        x.append(add_node(converted_files[file_path.as_posix()]['new'], metadata_col))
    else:
        failed_ref.append(file_path.as_posix())
        logging(f"{file_path.as_posix()} not found in SIARD-database")

## Skriver til SIARD-tabellen
root.write(temp_dir+ siard_table, pretty_print=True, encoding='utf-8')

## Legger til felter i metadata.xml
root_metadata = etree.parse(temp_dir + "header/metadata.xml")
root_metadata_column = root_metadata.xpath('/siard-metadata:siardArchive/siard-metadata:schemas/siard-metadata:schema/siard-metadata:tables/siard-metadata:table/siard-metadata:folder[text()="' + table + '"]/../siard-metadata:columns', namespaces=namespace)[0]
root_metadata.xpath('/siard-metadata:siardArchive/siard-metadata:schemas/siard-metadata:schema/siard-metadata:tables/siard-metadata:table/siard-metadata:folder[text()="' + table + '"]/../siard-metadata:columns', namespaces=namespace)[0]
root_metadata_column.append(add_node_metadataxml('filreferanse'))
root_metadata.write(temp_dir + "header/metadata.xml", pretty_print=True, encoding='utf-8')

## Legger til i xsden til tabellen
root_xsd = etree.parse(temp_dir + siard_table[:-3] + 'xsd')
root_xsd_element = root_xsd.xpath('//xml-schema:complexType[@name = "recordType"]/xml-schema:sequence', namespaces=namespace)[0]
root_xsd_element.append(add_node_xsd(metadata_col))
root_xsd.write(temp_dir + siard_table[:-3] + 'xsd', pretty_print=True, encoding='utf-8')

shutil.make_archive(f"{pathlib.Path(siard_filename).stem}_conv", 'zip', temp_dir)
shutil.move(f"{pathlib.Path(siard_filename).stem}_conv.zip", f"{pathlib.Path(siard_filename).stem}_conv.siard")

## Results file
pronom_data = json.load(open(f"{results_dir}pronom_stats.json"))

results_output = ''
results_output += f"Antall dokumenterfil i uttrekket: {len(converted_files)}\n"
results_output += f"Antall referanser i siardfilen: {siard_refs}\n"

file_no_ref_list = []
for x in converted_files:
    if converted_files[x]['siard'] == 'n':
        file_no_ref_list.append(converted_files[x]['new'])

if file_no_ref_list != []:
    results_output += '\n'
    results_output += f"Filer i uttrekket som mangler referanse fra SIARD ({len(file_no_ref_list)})\n"
    results_output += '\n'.join(file_no_ref_list)
    results_output += '\n'

if failed_ref != []:
    results_output += '\n'
    results_output += f"Referanser i SIARD som mangler filer ({len(failed_ref)})\n"
    results_output += '\n'.join(failed_ref)
    results_output += '\n'

antall_konverterte = 0
konverterte_pronom = {}
antall_ukonverterte = 0
ukonverterte_pronom = {}
antall_tomme = 0
tomme = []
timed = []

for x in converted_files:
    if converted_files[x]['conv'] == 'conv':
        antall_konverterte += 1
        if converted_files[x]['pronom'] in konverterte_pronom:
            konverterte_pronom[converted_files[x]['pronom']] += 1
        else:
            konverterte_pronom[converted_files[x]['pronom']] = 1
    if converted_files[x]['conv'] == 'nconv':
        antall_ukonverterte += 1
        if converted_files[x]['pronom'] in ukonverterte_pronom:
            ukonverterte_pronom[converted_files[x]['pronom']] += 1
        else:
            ukonverterte_pronom[converted_files[x]['pronom']] = 1
    if converted_files[x]['conv'] == 'time':
        timed.append(converted_files[x]['new'])
    if converted_files[x]['empty'] == 'y':
        tomme.append(converted_files[x]['new'])


results_output += '\n'
results_output += f"Konverterte filer ({antall_konverterte})\n"
for key, value in sorted(konverterte_pronom.items(), key=lambda item: item[1], reverse = True):
    results_output += f"{str(value).rjust(10, ' ')}\t{key.ljust(12, ' ')}\t{pronom_data['names'][key]}\n"
results_output += '\n'
results_output += f"Ukonverterte filer ({antall_ukonverterte})\n"
for key, value in sorted(ukonverterte_pronom.items(), key=lambda item: item[1], reverse = True):
    results_output += f"{str(value).rjust(10, ' ')}\t{key.ljust(12, ' ')}\t{pronom_data['names'][key]}\n"
results_output += '\n'
if timed != []:
    results_output += f"Filer som timet ut i prosessen ({len(timed)})\n"
    results_output += '\n'.join(timed)
results_output += '\n'
results_output += '\n'
if tomme != []:
    results_output += f"Filer som var tomme (0 byte) ({len(tomme)})\n"
    results_output += '\n'.join(tomme)

results_file = open(f'{results_dir}results.txt','w')
results_file.write(results_output)
results_file.close()

logging(f"Done")
log_file.close()

