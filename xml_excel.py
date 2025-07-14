import xml.etree.ElementTree as ET
import re
from io import BytesIO
from datetime import datetime
from os.path import exists, splitext
from zipfile import ZipFile

# Read the file, then close
class Workbook:

    def __init__(self, fp: str):
        
        self._validate_path(fp)

        self.fp = fp
        self._zip = None


    # With/open support
    def __enter__(self):

        self._zip = ZipFile(self.fp, "r")

        self._prepare_excel_file()

        self._set_props()

        return self
    

    def __exit__(self, exc_type, exc_value, traceback):
 
        self._zip.close()
        self._zip = None

    
    # Manual file handling
    def open(self):

        return self.__enter__()
   

    def close(self):

        if self._zip is not None:

            self._zip.close()
            self._zip = None

        
    def _validate_path(self, fp: str):

        if not isinstance(fp, str):
            raise UserWarning("The file path is not str.")
        
        _, extension = splitext(fp)
        if extension not in (".xlsx", ".xlsm"):
            raise UserWarning(f"The {extension} file extension is not supported.")
        
        if not exists(fp):
            raise UserWarning("The file could not be found.")


    def _prepare_excel_file(self):
        
        required_folders = ['docProps', 'xl', 'xl/theme', 'xl/worksheets']
        required_files = ['docProps/core.xml', 'docProps/app.xml']

        self._folders = self.list_folders()

        for folder in required_folders:
            if folder not in self._folders:

                raise UserWarning(f"File corrupted: the {folder} folder is not found in the excel file.")

        current_files = [x.filename for x in self._zip.filelist]

        for file in required_files:
            if file not in current_files:
                raise UserWarning(f"File corrupted: the {file} xml file is not found in the excel file.")

        # Create the sharedstrings.xml file if the excel file is empty
        self.add_sharedstrings()


    def list_folders(self) -> list[str]:

        folders = []

        for full_path in self._zip.filelist:

            split_path = full_path.filename.split("/")

            parent = ""
            for i in range(len(split_path) - 1):

                folder = split_path[i] if parent == "" else parent + "/" + split_path[i]

                if folder not in folders:

                    folders.append(folder)

                parent = folder

        return folders
    

    def add_sharedstrings(self):
        
        fn = "xl/sharedStrings.xml"

        # Create the file in case it does not exist
        if fn not in [x.filename for x in self._zip.filelist]:

            xml_content = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                             <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
                                <si>
                                    <t>temp_string</t>
                                </si>
                             </sst>'''

            xml_content.encode('utf-8')


            self._zip.close()

            with ZipFile(self.fp, 'a') as file:
                file.writestr(fn, xml_content)

            self._zip = ZipFile(self.fp, 'r')

            # Add the sharedStrings data into the relationships file so Excel is able to communicate with it
            rel_stream = get_stream(self._zip, "xl/_rels/workbook.xml.rels")
            rel_root = ET.ElementTree(ET.fromstring(rel_stream)).getroot()
            ns = rel_root.tag.split('}')[0][1:]

            new_id = self._next_id(rel_stream)

            rel_node = ET.SubElement(rel_root, f"{{{ns}}}Relationship")
            rel_node.set('Id', new_id)
            rel_node.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings')
            rel_node.set('Target', 'sharedStrings.xml')

            # Refresh the content types.xml
            content_stream = get_stream(self._zip, '[Content_Types].xml')
            content_root = ET.ElementTree(ET.fromstring(content_stream)).getroot()
            content_ns = content_root.tag.split('}')[0][1:]
            content_node = ET.SubElement(content_root, f"{{{content_ns}}}Override")
            content_node.set('PartName', '/xl/sharedStrings.xml')
            content_node.set('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')

            # Save all modifications
            self._save_xml(
                ["xl/_rels/workbook.xml.rels", "[Content_Types].xml"],
                [ET.tostring(rel_root, encoding='utf-8'), ET.tostring(content_root, encoding='utf-8')]
            )


    def _next_id(self, stream: str) -> str:
        """Get next available ID in .rels file.

        Args:
            stream (str): Content of .rels file.
        Returns:
            str: Next available ID.
        """
        tree = ET.ElementTree(ET.fromstring(stream))

        relationships = tree.getroot().findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")

        ids = {int(rel.attrib['Id'].replace('rId', '')) for rel in relationships}

        return f"rId{max(ids, default=0) + 1}"
    

    def _save_xml(self, paths: list, streams: list):

        # List every file from the excel into the memory
        temp_buffer = BytesIO()

        with ZipFile(temp_buffer, 'w') as zip_buffer:
            for item in self._zip.infolist():
                content = self._zip.read(item.filename)
                zip_buffer.writestr(item, streams[paths.index(item.filename)] if item.filename in paths else content)

        # Put back every element
        with ZipFile(self.fp, 'w') as file:
            with ZipFile(temp_buffer, 'r') as temp:
                for item in temp.infolist():
                    file.writestr(item, temp.read(item.filename))

        self._zip = ZipFile(self.fp, 'r')

    
    def _set_props(self):


        self.modification_date = self.return_date()

        self.version = self.return_version()

        self.sheets = self.return_sheets()

        self.sheet_metadata = self.return_metadata()


    def return_date(self) -> datetime | None:

        '''Returns the modification date of the file.'''

        date = None

        if self._zip is not None:

            date = process_xml(self._zip, "docProps/core.xml", ".//dcterms:modified", "text")[0]
            date = datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ")

        return date


    def return_version(self) -> str | None:

        '''Returns the Excel version the file was created in.'''

        version = None

        if self._zip is not None:

            version = process_xml(self._zip, "docProps/app.xml", ".//default:AppVersion", "text")[0]

        return version


    def return_sheets(self) -> dict:

        sheet_dict = {}

        if self._zip is not None:

            data = process_xml(self._zip, 'xl/workbook.xml', ".//default:sheets/default:sheet", "value")

            for elem in data:

                sheet_dict[elem['name']] = elem[list(elem.keys())[-1]]

        return sheet_dict


    def return_metadata(self) -> dict:

        '''Create a dictionary for sheet_ids and sheet_paths.'''

        metadata = {}

        if self._zip is not None:

            stream = get_stream(self._zip, "xl/_rels/workbook.xml.rels")

            tree = ET.ElementTree(ET.fromstring(stream))
            relationships = tree.getroot().findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")

            for relation in relationships:

                metadata[relation.attrib['Id']] = relation.attrib['Target']

        return metadata
    

    def read_sheet(self, sheet_id: str | int, headers: list | None=None) -> list[list[str]]:

        if self._zip is not None:

            sheet_id = self._get_sheet_id(sheet_id)

            # Decide which file to read from
            key = self.sheets[sheet_id]
            path = f"xl/{self.sheet_metadata[key]}"

            # Get the general file which stores every value accross each sheet as a dict
            values = process_xml(self._zip, "xl/sharedStrings.xml", ".//default:si/default:t", "text")
            values = {index: value for index, value in enumerate(values)}

            # Get the coordinates for the active range
            end_point = self._get_dimension(path)
            end_point = self._translate_end_point(end_point)

            # Generate empty table to populate with values
            self.sheet = self.Worksheet(self, end_point, key, path)

            # Populate values
            self.sheet.populate_table(values, self._translate_end_point)

            # Replace the headers
            if headers is not None:
                if len(headers) == len(self.sheet.table[0]):
                    for index, header in enumerate(headers):
                        self.sheet.table[0][index] = header

            return self.sheet.table

 

 

    def _get_dimension(self, path: str) -> str:

        stream = get_stream(self._zip, path)

        namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

        result = get_xml_value(stream, ".//default:dimension", namespaces, "value")[0]

        result = list(result.values())[0]

        result = result.split(":")

        return result[-1]

 
    def upload_sheet(self, sheet_id: str | int, new_table: list[list]):

        '''Upload a table to the target sheet, replacing all values found in there.

        sheet_id: (int | str) - identififer for the sheet.

        new_table: (list) - data to upload.'''

 

        if self._zip is not None:

            # Get the sheet
            sheet_id = self._get_sheet_id(sheet_id)

            # Decide which file to read from
            key = self.sheets[sheet_id]
            path = f"xl/{self.sheet_metadata[key]}"


            # Gather the data from the XML file
            stream = get_stream(self._zip, path)
            namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))
    
            pattern = "<.+?>"
            header = ""; counter = 0

            for x in re.finditer(pattern, stream):
                header += x.group(); counter += 1
                if counter == 2: break

            # Get the tree
            tree = ET.ElementTree(ET.fromstring(stream))
            root = tree.getroot()

            # Analyze the new table
            end_point = "A1" if not new_table else self._translate_coords(len(new_table), len(new_table[0]))

            # Remove the content from the
            root = self._remove_child_nodes(root, namespaces, ".//default:sheetData")

            # Insert the table
            stream = self._insert_xml_node(root, namespaces, new_table, end_point)

 
            # Issue - ET library discards unused namespace --> get original header and insert it back to the stringified xml content
            stream = re.sub(pattern, header, stream, 1).encode('utf8')

            # Refresh the xml file
            self._save_xml([path], [stream])

            # After every modification refresh the current_sheet
            self.sheet = self.Worksheet(self, (len(new_table), len(new_table[0])), key, path)
            self.sheet.table = new_table

 
    def _translate_coords(self, row_count: int, col_count:int) -> str:

        range = ""

        while col_count > 0:

            col_count -= 1

            range = chr(ord('A') + col_count % 26) + range

            col_count //= 26

        range += str(row_count)

        return range
    

    def _remove_child_nodes(self, root: ET.Element, namespaces: dict[str, str], node_name: str) -> ET.Element:

        node = root.findall(node_name, namespaces)[0]
        for child in list(node):
            node.remove(child)
        return root
    

    def _insert_xml_node(self, root: ET.ElementTree, namespaces: dict, table: list, end_point: str) -> str:

        # Create the sharedstrings.xml file if not found
        self.add_sharedstrings()

        # Gather sharedStrings data
        stream = get_stream(self._zip, "xl/sharedStrings.xml")
        str_xml_ns = gather_namespaces(BytesIO(stream.encode('utf-8')))


        # Create the hierarchy object for the sharedStrings.xml
        str_xml = ET.ElementTree(ET.fromstring(stream))
        str_xml_root = str_xml.getroot()
        str_xml_count = int(str_xml_root.attrib.get('count', 0))
        original_count = str_xml_count

        # First modify the endpoint
        dim = root.findall(".//default:dimension", namespaces)[0]
        dim.set('ref', f"A1:{end_point}")

        sheet_data = root.findall(".//default:sheetData", namespaces)[0]
        base_ns = namespaces['default']

        # Add every value to the files
        for i, row in enumerate(table, 1):

            row_node = ET.SubElement(sheet_data, ET.QName(base_ns, 'row'))
            row_node.set('r', str(i))
            row_node.set('spans', f'1:{len(row)}')
            row_node.set('x14ac:dyDescent', '0.25')

            for j, element in enumerate(row, 1):

                cell_node = ET.SubElement(row_node, ET.QName(base_ns, 'c'))
                cell_node.set('r', self._translate_coords(i, j))
 
                value_node = ET.SubElement(cell_node, ET.QName(base_ns, 'v'))

                # Handle string elements, they are stored in the sharedStrings.xml, the sheet file only contains the reference to this element
                if isinstance(element, str):

                    cell_node.set('t', 's')

                    found, index = self._is_string_used(str_xml_root, str_xml_ns, element)

                    if found:
                        value_node.text = str(index)

                    else:

                        # Instead of the value we insert the index to the sharedStrings.xml element
                        value_node.text = str(str_xml_count)

                        str_xml_count += 1

                        # Add the original value to sharedstrings.xml
                        str_xml_node = ET.SubElement(str_xml_root, ET.QName(base_ns, 'si'))
                        str_xml_val = ET.SubElement(str_xml_node, ET.QName(base_ns, 't'))
                        str_xml_val.text = element

                else:
                    value_node.text = str(element)

 

        # Upload back the sheet
        if original_count != str_xml_count:

            str_xml_root.set('count', str(str_xml_count))
            str_xml_root.set('uniqueCount', str(str_xml_count))

        str_xml_str = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n{ET.tostring(str_xml_root, encoding="unicode")}'
        self._save_xml(["xl/sharedStrings.xml"], [str_xml_str.encode('utf-8')])
 
        return ET.tostring(root, encoding='unicode')


    def _translate_end_point(self, end_point: str) -> tuple[int]:

        column = re.search("[A-Z]+", end_point).group(0)

        row_index = int(end_point.replace(column, ""))

        column_index = 0

        for char in column:

            column_index = column_index * 26 + (ord(char) - ord("A")) + 1

        return row_index, column_index
 

    def _get_sheet_id(self, sheet_id: int | str) -> str:

        sheets = list(self.sheets.keys())

        # Handle invalid id-s
        if isinstance(sheet_id, str):
            if sheet_id not in sheets:
                raise UserWarning(f"Invalid sheet name: {sheet_id}")

        else:
            if sheet_id > len(sheets):
                raise UserWarning(f"Invalid sheet index: {sheet_id}")
            
            sheet_id = sheets[sheet_id]

        return sheet_id
    

    def _is_string_used(self, str_xml_root: ET.ElementTree, namespaces: dict, value_to_search: str) -> tuple[bool, int]:

        items = {x.text: i for i, x in enumerate(str_xml_root.findall(".//default:si/default:t", namespaces)) if x.text}

        if value_to_search in items:

            return True, list(items.keys()).index(value_to_search)

        else:

            return False, len(items)
    

    class Worksheet:


        def __init__(self, parent: object, end_point: tuple[int], rid: str, path: str):


            self.parent = parent
            self.ID = rid
            self.path = path
            self.row_count, self.column_count = end_point
            self.table = self._resize_table()


        def _resize_table(self):

            table = [["" for _ in range(self.column_count)] for _ in range(self.row_count)]

            return table


        def populate_table(self, values: dict, func: object):

            data_pairs, direct_values = self._process_sheet()

            # Handle string values
            for cell_id, sheet_id in data_pairs.items():
                row, col = self.parent._translate_end_point(cell_id)
                self.table[row - 1][col - 1] = values[sheet_id]

            # Handle other values
            for cell_id, value in direct_values.items():
                row, col = self.parent._translate_end_point(cell_id)
                self.table[row - 1][col - 1] = value


        def _process_sheet(self) -> tuple[dict, dict]:

            data_pairs, values = {}, {}

            stream = get_stream(self.parent._zip, self.path)

            namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

            for node in get_xml_value(stream, ".//default:sheetData/default:row/default:c", namespaces):

                is_string = node.attrib.get('t')

                position = node.attrib['r']

                value = node.find(".//default:v", namespaces)

                if value is not None:

                    value = value.text

                    if is_string == 's':
                        data_pairs[position] = int(value)
                    else:
                        values[position] = value

            return data_pairs, values


def process_xml(zip: ZipFile, filename: str, search_str: str, attrib_type: str="") -> list[str]:

    stream = get_stream(zip, filename)

    namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

    values = get_xml_value(stream, search_str, namespaces, attrib_type)

    return values


def get_xml_value(stream: str, search_str: str, namespaces: dict, attrib_type: str="") -> None | list[ET.Element] | list[dict] | str:


    output = None

    tree = ET.ElementTree(ET.fromstring(stream))

    root = tree.getroot()

    output = root.findall(search_str, namespaces)

    if attrib_type == "text":
        output = [elem.text for elem in output]

    if attrib_type == "value":

        tmp_elem = [elem.attrib for elem in output]

        output = []
        for elem in tmp_elem:

            output.append({k: v for k, v in elem.items()})

    return output

 



def gather_namespaces(file):

    namespaces = {}

    for event, elem in ET.iterparse(file, ("start", "start-ns")):

        if event == "start-ns":

            key, value = elem
            key = "default" if key == "" else key
            namespaces[key] = value

            ET.register_namespace(key if key != "default" else '', value)

        if event == "start":
            break

    return namespaces


def get_stream(comp: ZipFile, filename: str) -> str:

    with comp.open(filename) as file:

        return file.read().decode('utf-8')
    

with Workbook(r"E:\asd.xlsx") as f:
    print(f.read_sheet(0))
    t = [["haha", "bb"]]
    f.upload_sheet(0, t)
    print(f.sheet.table)

