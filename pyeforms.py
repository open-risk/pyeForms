# (c) 2024 - 2025 Open Risk (www.openriskmanagement.com), all rights reserved
#
# pyeForms is licensed under the Apache 2.0 license a copy of which is included
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

    name = sys.argv[1]

    # Create a new Excel file and add a worksheet.
    filename = os.path.basename(name)
    xlsx_name = filename.split('.')[0] + '.xlsx'
    print(xlsx_name)
    workbook = xlsxwriter.Workbook('examples/' + xlsx_name)
    worksheet = workbook.add_worksheet('About')
    worksheet.write(0, 0, 'pyeForms demo')

    #
    # Create XLSX Data Sheets for all unique form sections
    # NB: some may not be populated in an given notice
    #
    field_file2 = open('reference/fields.csv')
    csvreader3 = csv.reader(field_file2, delimiter=',')
    # Find the unique sections
    sections = set([row[10].capitalize() for row in csvreader3])
    ws_dict = {}
    for section in sections:
        worksheet = workbook.add_worksheet(section)
        ws_dict[section] = 0  # row index per sheet
    field_file2.close()

    #
    # Open and parse an XML Notice file
    #
    notice_file = open(name, 'rb')
    notice_xml = notice_file.read()
    xml_root = etree.XML(notice_xml)
    nsa = xml_root.nsmap
    namespaces = {k: v for k, v in nsa.items() if k is not None}
    # Set the notice type
    notice_type = xml_root.tag.split('}')[1]

    #
    # Create Node Data Sheet (TEMP)
    #

    worksheet = workbook.add_worksheet('Node Data')

    # Open and parse Node file
    node_file = open('reference/nodes.csv')
    csvreader1 = csv.reader(node_file, delimiter=',')
    i = 0
    k = 0
    for row in csvreader1:
        xpathRelative = '//' + row[2]
        r = xml_root.xpath(xpathRelative, namespaces=namespaces)
        if len(r) > 0 and i != 241:
            tag = r[0].tag.split('}')[1]
            # print(i, tag, len(r), row[3])
            j = 0
            for elem in r:
                # print('>>> ', j, elem)
                worksheet.write(k, 0, tag)
                worksheet.write(k, 1, row[3])
                j += 1
                k += 1
        else:
            pass
        i += 1
    notice_file.close()

    #
    # Create Field Data Sheets
    #

    # worksheet = workbook.add_worksheet('Field Data')

    # Open and parse Fields file
    field_file = open('reference/fields.csv')
    csvreader2 = csv.reader(field_file, delimiter=',')
    i = 1
    k = 0
    # Iterate over the fields dictionary and find all corresponding elements in the notice XML
    for row in csvreader2:
        xpathAbsolute = '//' + row[9][3:]
        r = xml_root.xpath(xpathAbsolute, namespaces=namespaces)
        if len(r) > 0:
            if hasattr(r[0], 'text'):
                value = r[0].text.strip()
                # print(i,  value, row[5])
            else:
                value = r[0]
                # print(i, value, row[5])
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
        print(elem, ws_dict[elem])
        if ws_dict[elem] == 0:
            worksheet.hide()
            print('Hiding: ', elem)

    workbook.close()
