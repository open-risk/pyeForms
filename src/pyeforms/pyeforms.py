# (c) 2024 - 2026 Open Risk (www.openriskmanagement.com), all rights reserved
#
# pyeForms is licensed under the MIT license a copy of which is included
# in the source distribution of pyeForms. This is notwithstanding any licenses of
# third-party software included in this distribution. You may not use this file except in
# compliance with the License.
#
# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
# either express or implied. See the License for the specific language governing permissions and
# limitations under the License.

import os
import sys
import csv

from lxml import etree
import xlsxwriter

if __name__ == "__main__":

    # Read the input filename from the command line
    name = sys.argv[1]
    filename = os.path.basename(name)

    # Create a new XLSX file and add a worksheet "About".
    xlsx_name = filename.split('.')[0] + '.xlsx'
    workbook = xlsxwriter.Workbook('examples/' + xlsx_name)
    worksheet = workbook.add_worksheet('About')
    worksheet.write(0, 0, 'pyeForms demo')

    # Create XLSX Data Sheets for all unique **form sections**
    # NB: some may not be populated for a given TED/eForms notice
    field_file = open('reference/fields.csv')
    csv1 = csv.reader(field_file, delimiter=',')
    # Find the unique sections (C11)
    sections = set([row[10].capitalize() for row in csv1])
    ws_dict = {}
    for section in sections:
        worksheet = workbook.add_worksheet(section)
        ws_dict[section] = 0  # row index per sheet
    field_file.close()

    # Open and parse an XML eForms Notice file
    notice_file = open(name, 'rb')
    notice_xml = notice_file.read()
    xml_root = etree.XML(notice_xml)
    nsa = xml_root.nsmap
    namespaces = {k: v for k, v in nsa.items() if k is not None}
    # Set the notice type
    notice_type = xml_root.tag.split('}')[1]

    # Create Node Data Sheet (TEMP)
    worksheet = workbook.add_worksheet('Node Data')
    # Open and parse Node file
    node_file = open('reference/nodes.csv')
    csv2 = csv.reader(node_file, delimiter=',')
    i = 0
    k = 0
    for row in csv2:
        xpathRelative = '//' + row[2]
        r = xml_root.xpath(xpathRelative, namespaces=namespaces)
        if len(r) > 0 and i != 241:
            tag = r[0].tag.split('}')[1]
            j = 0
            for elem in r:
                worksheet.write(k, 0, tag)
                worksheet.write(k, 1, row[3])
                j += 1
                k += 1
        else:
            pass
        i += 1
    notice_file.close()

    # Create Field Data Sheets
    # Open and parse Fields file
    field_file = open('reference/fields.csv')
    csv3 = csv.reader(field_file, delimiter=',')
    i = 1
    k = 0
    # Iterate over fields and find all corresponding elements in the notice XML
    for row in csv3:
        xpathAbsolute = '//' + row[9][3:]
        r = xml_root.xpath(xpathAbsolute, namespaces=namespaces)
        if len(r) > 0:
            if hasattr(r[0], 'text'):
                value = r[0].text.strip()
            else:
                value = r[0]
            # write the found element to the corresponding sheet / section
            section = row[10].capitalize()
            worksheet = workbook.get_worksheet_by_name(section)
            k = ws_dict[section]
            worksheet.write(k, 0, row[3])
            worksheet.write(k, 1, value)
            worksheet.write(k, 2, row[5])
            ws_dict[section] += 1
        i += 1
    field_file.close()

    # Wrap things up
    for elem in ws_dict:
        worksheet = workbook.get_worksheet_by_name(elem)
        worksheet.autofit()
        if ws_dict[elem] == 0:
            worksheet.hide()
            print('Hiding: ', elem)

    workbook.close()