from datetime import datetime
from io import BytesIO
from os.path import exists, splitext
from zipfile import ZipFile
import re
import xml.etree.ElementTree as ET


class Workbook:

    def __init__(self, fp: str):
        """Initializes the Workbook instance.

        Args:
            fp (str): File path to the workbook to open.
        """
        
        self._validate_path(fp)

        self.fp = fp
        self._zip = None


    # With/open support
    def __enter__(self):

        self._zip = ZipFile(self.fp, "r")

        self._file_integrity_assessment()

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

    
    # Init methods
    def _validate_path(self, fp: str):
        """Checks whether the filepath is supported.

        Args:
            fp (str): Path of the excel file.
        Returns:
            None
        """

        if not isinstance(fp, str):
            raise UserWarning("The file path is not str.")
        
        _, extension = splitext(fp)
        if extension not in (".xlsx", ".xlsm"):
            raise UserWarning(f"The {extension} file extension is not supported.")
        
        if not exists(fp):
            raise UserWarning("The file could not be found.")


    # File modification
    def _save_xml(self, paths: list, streams: list):
        """Saves the modifications done on the excel file.

        Args:
            paths (list[str]): List of files to modify.
            streams (list): List of new xml files in bytes.
        Returns:
            None
        """

        # TODO - replace function, to use the harddrive instead of memory

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


    # Methods after open is ran
    def _file_integrity_assessment(self):
        """Checks for excel file integrity.

        Returns:
            None
        """
        
        required_folders = ['docProps', 'xl', 'xl/theme', 'xl/worksheets']
        required_files = ['docProps/core.xml', 'docProps/app.xml']

        self._folders = self._list_folders()

        for folder in required_folders:
            if folder not in self._folders:

                raise UserWarning(f"File corrupted: the {folder} folder is not found in the excel file.")

        current_files = [x.filename for x in self._zip.filelist]

        for file in required_files:
            if file not in current_files:
                raise UserWarning(f"File corrupted: the {file} xml file is not found in the excel file.")

        # Create the sharedstrings.xml file if the excel file is empty
        self._add_sharedstrings()


    def _list_folders(self) -> list[str]:
        """Returns every folder found in the compressed excel file.

        Returns:
            list[str]: List of folders.
        """

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
    

    def _add_sharedstrings(self):
        """Creates the sharedStrings.xml file in case it is missing. The relationships are also added to other xml files.

        Returns:
            None
        """
        
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

    
    def _set_props(self):
        """Sets basic data from the excel file to the class instance."""

        self.modification_date = self.return_date()

        self.version = self.return_version()

        self.sheet_ids = self.return_sheets()

        self.sheet_metadata = self.return_metadata()

        self.sheets = self.prepare_sheet_container()


    def return_date(self) -> datetime | None:
        """Returns the modification date of the file.

        Returns:
            datetime: for the date.
            None: if class instance is reached after file is closed.
        """

        date = None

        if self._zip is not None:

            date = process_xml(self._zip, "docProps/core.xml", ".//dcterms:modified", "text")[0]
            date = datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ")

        return date


    def return_version(self) -> str | None:
        """Returns the excel version.

        Returns:
            str: version number.
            None: if class instance is reached after file is closed.
        """

        version = None

        if self._zip is not None:

            version = process_xml(self._zip, "docProps/app.xml", ".//default:AppVersion", "text")[0]

        return version


    def return_sheets(self) -> dict:
        """Returns every sheet.

        Returns:
            dict: A dictionary where each key is a sheet name (str) and the corresponding value
                  is the sheet's associated attribute value extracted from the XML. 
                  For example, {'Sheet1': '1', 'Sheet2': '2', ...}.
                  If no sheets are found or the ZIP archive is not loaded, returns an empty dictionary
        """

        sheet_dict = {}

        if self._zip is not None:

            data = process_xml(self._zip, 'xl/workbook.xml', ".//default:sheets/default:sheet", "value")

            for elem in data:

                sheet_dict[elem['name']] = elem[list(elem.keys())[-1]]

        return sheet_dict


    def return_metadata(self) -> dict:
        """Returns the metadata relationships from the workbook's .rels file.

        Returns:
            dict: A dictionary where each key is a relationship ID (str) and the value is the
                target path (str) of that relationship. For example:
                {'rId1': 'worksheets/sheet1.xml', 'rId2': 'worksheets/sheet2.xml', ...}.
                Returns an empty dictionary if the .rels file is missing or the ZIP archive is not loaded.
        """

        metadata = {}

        if self._zip is not None:

            stream = get_stream(self._zip, "xl/_rels/workbook.xml.rels")

            tree = ET.ElementTree(ET.fromstring(stream))
            relationships = tree.getroot().findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")

            for relation in relationships:

                metadata[relation.attrib['Id']] = relation.attrib['Target']

        return metadata
    

    def prepare_sheet_container(self) -> dict:
        """Prepares a dictionary to later store each sheet found in the workbook.

        Returns:
            dict: A dictionary where each key is a sheet name (str) and the corresponding value
                  is an empty list.
        """

        return {sheet_id: None for sheet_id in self.sheet_ids}


    # General methods to handle excel logic
    def _translate_coords(self, row: int, col: int) -> str:
        """Translates the column and row indexes into a cell ID ie.: A18... (column + row)

        Args:
            row (int): Row number.
            col (int): Col number.
        Returns:
            str: Cell id.
        """

        # Avoid the confusion, as lists and excel worksheets start the ir indexing at 0 and 1
        if 0 in [row, col]:
            raise UserWarning("Row and column indices must be greater than zero (1-based indexing).")

        range = ""

        while col > 0:

            col -= 1

            range = chr(ord('A') + col % 26) + range

            col //= 26

        range += str(row)

        return range
    

    def _translate_end_point(self, end_point: str) -> tuple[int]:
        """Converts an Excel-style cell reference (e.g. 'C12') into a tuple of (row, column) indexes.
        Note - the row does not count with the 0-based indexing of python.

        Args:
            end_point (str): Cell reference string in Excel format (letters + digits), e.g. 'A1', 'BC23'.

        Returns:
            tuple[int, int]: A tuple where the first element is the row number (int),
                             and the second element is the column number (int).
        """

        column = re.search("[A-Z]+", end_point).group(0)

        row_index = int(end_point.replace(column, ""))

        column_index = 0

        for char in column:

            column_index = column_index * 26 + (ord(char) - ord("A")) + 1

        return row_index, column_index
 

    def _get_sheet_id(self, sheet_id: int | str) -> str:
        """Resolves a sheet identifier given as either a sheet name (str) or an index (int) to a valid sheet name.

        Args:
            sheet_id (int | str): The sheet identifier, either as a zero-based index (int) or a sheet name (str).

        Returns:
            str: The validated sheet name corresponding to the provided identifier.
        """

        sheets = list(self.sheet_ids.keys())

        # Handle invalid id-s
        if isinstance(sheet_id, str):
            if sheet_id not in sheets:
                raise UserWarning(f"Invalid sheet name: {sheet_id}")

        else:
            if sheet_id > len(sheets):
                raise UserWarning(f"Invalid sheet index: {sheet_id}")
            
            sheet_id = sheets[sheet_id]

        return sheet_id
    

    # Read methods
    def read_all(self):
        """Reads the content of all sheets in the workbook and returns them as a list of 2D string lists.

        Returns:
            list[list[list[str]]]: A list where each element is a 2D list representing one sheet.
                                   Each 2D list contains rows, and each row is a list of string cell values.
        """


        if self._zip is None:
            return []

        return [self.read_sheet(i) for i in range(len(self.sheets))]



    def read_sheet(self, sheet_id: str | int, headers: list | None=None) -> list[list[str]]:
        """Reads the content of a specified sheet from the workbook and returns it as a 2D list of strings.

        Args:
            sheet_id (str | int): Name or index of the sheet.
            headers (list | None, optional): A list of header strings to replace the first row's headers.
                                             If None, original headers are kept. Defaults to None.

        Returns:
            list[list[str]]: A 2D list representing the sheet's rows and columns,
                             where each inner list is a row of string cell values.
        """

        if self._zip is not None:

            sheet_id = self._get_sheet_id(sheet_id)

            # Decide which file to read from
            key = self.sheet_ids[sheet_id]
            path = f"xl/{self.sheet_metadata[key]}"

            # Get the general file which stores every value accross each sheet as a dict
            values = process_xml(self._zip, "xl/sharedStrings.xml", ".//default:si/default:t", "text")
            values = {index: value for index, value in enumerate(values)}

            # Get the coordinates for the active range
            end_point = self._get_dimension(path)
            end_point = self._translate_end_point(end_point) # Don't subtract 1 --> end point is used as the upper bound of "range()"

            # Generate empty table to populate with values
            sheet = self.Worksheet(self, end_point, key, path)
            self.sheets[sheet_id] = sheet

            # Populate values
            sheet.populate_table(values)

            # Replace the headers
            if headers is not None:
                if len(headers) == len(self.sheet.table[0]):
                    for index, header in enumerate(headers):
                        self.sheet.table[0][index] = header

            return sheet.table


    def _get_dimension(self, path: str) -> str:
        """Returns the last used cell ID of the given sheet ie.: A8, C95, ZA51... (column + row)

        Args:
            path (str): XML path for the sheet.
        Returns:
            str: Identifier for the cell.
        """

        stream = get_stream(self._zip, path)

        namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

        result = get_xml_value(stream, ".//default:dimension", namespaces, "value")[0]

        result = list(result.values())[0]

        result = result.split(":")

        return result[-1]

 
    # Write methods
    def upload_sheet(self, sheet_id: str | int, new_table: list[list]):
        """Replaces the content of the sheet.

        Args:
            sheet_id (str | int): Name or index of the sheet.
            new_table (list[list]) - 2D list of the new content
        Returns:
            None
        """

        if self._zip is not None:

            # Get the sheet
            sheet_id = self._get_sheet_id(sheet_id)

            # Decide which file to read from
            key = self.sheet_ids[sheet_id]
            path = f"xl/{self.sheet_metadata[key]}"

            # Gather the data from the XML file
            stream = get_stream(self._zip, path)
            namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

            # Get the tree
            tree = ET.ElementTree(ET.fromstring(stream))
            root = tree.getroot()

            # Analyze the new table
            end_point = "A1" if not new_table else self._translate_coords(len(new_table), len(new_table[0]))

            # Remove the content from the
            root = remove_child_nodes(root, namespaces, ".//default:sheetData")

            # Insert the table
            stream = self._insert_new_values_to_xml(root, namespaces, new_table, end_point)

            # Refresh the xml file
            self._save_xml([path], [stream.encode('utf8')])

            # After every modification refresh the current_sheet
            sheet = self.Worksheet(self, (len(new_table), len(new_table[0])), key, path)
            sheet.table = new_table
            self.sheets[sheet_id] = sheet
            

    def _insert_new_values_to_xml(self, root: ET.ElementTree, namespaces: dict, table: list[list], end_point: str) -> str:
        """Updates the sharedStrings xml and the XML file for the given sheet with the new content.

        Args:
            root (ET.Element): The XML root of the sheet.
            namespaces (dict[str, str]): Namespace mappings used for XPath queries.
            table (list[list]): 2D list to upload.
            end_point (str): Last cell of the used range.

        Returns:
            str: The modified root.
        """

        # Create the sharedstrings.xml file if not found
        self._add_sharedstrings()

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

        str_xml_str = ET.tostring(str_xml_root, encoding='unicode')
        self._save_xml(["xl/sharedStrings.xml"], [str_xml_str.encode('utf-8')])
 
        return ET.tostring(root, encoding='unicode')
    

    def _is_string_used(self, str_xml_root: ET.ElementTree, namespaces: dict, value_to_search: str) -> tuple[bool, int]:
        """Returns whether the string is already part of the sharedStrings.xml and its position.

        Args:
            str_xml_root (ET.ElementTree): Parsed XML tree of sharedStrings.xml.
            namespaces (dict): Namespace mapping used for XML queries.
            value_to_search (str): The string value to look for in the shared strings.

        Returns:
            tuple[bool, int]: A tuple where the first element is True if the string is found,
                              False otherwise; the second element is the index of the string 
                              if found, or the total number of strings in order to continue.
        """

        items = {x.text: i for i, x in enumerate(str_xml_root.findall(".//default:si/default:t", namespaces)) if x.text}

        if value_to_search in items:
            return True, list(items.keys()).index(value_to_search)

        else:
            return False, len(items)
    

    class Worksheet:

        def __init__(self, parent: object, end_point: tuple[int], rid: str, path: str):
            """Initializes a Worksheet instance.

            Args:
                parent (object): Reference to the parent Workbook.
                end_point (tuple[int, int]): Tuple of excel based (row_count, column_count) defining the size of the worksheet.
                rid (str): Relationship ID or unique identifier for the worksheet.
                path (str): Full name of the worksheet XML file.
            """

            self.parent = parent
            self.ID = rid
            self.path = path
            self.table = self._resize_table(*end_point)


        def _resize_table(self, row: int, col: int):
            """Prepares an empty table to store the sheet content.

            Args:
                row (int): Row count of the table (starting from 1).
                col (int): Column count of the table (starting from 1).
            Returns:
                None
            """

            table = [["" for _ in range(col)] for _ in range(row)]

            return table


        def populate_table(self, values: dict):
            """Fills the table with data.

            Args:
                values (dict): Mapping of shared string IDs to their corresponding string values.
            Returns:
                None
            """

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
            """Parses the worksheet XML to extract cell values, separating shared string references and direct values.

            Returns:
                tuple[dict[str, int], dict[str, str]]:
                    - First dict maps cell positions (e.g. 'A1') to shared string IDs (int).
                    - Second dict maps cell positions to direct string values (non-shared strings).
            """

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


# General XML methods
def process_xml(zip: ZipFile, filename: str, search_str: str, attrib_type: str="") -> list[str]:
    """Extracts and returns values from an XML file within a ZIP archive based on a given XPath and attribute type.

    Args:
        zip (ZipFile): Open ZipFile object containing the XML file.
        filename (str): Path to the XML file inside the ZIP archive.
        search_str (str): XPath expression to locate the desired XML elements.
        attrib_type (str, optional): Attribute or element name to extract text from. Defaults to "".

    Returns:
        list[str]: List of extracted string values matching the search criteria.
    """

    stream = get_stream(zip, filename)

    namespaces = gather_namespaces(BytesIO(stream.encode('utf-8')))

    values = get_xml_value(stream, search_str, namespaces, attrib_type)

    return values


def get_xml_value(stream: str, search_str: str, namespaces: dict, attrib_type: str="") -> None | list[ET.Element] | list[dict] | str:
    """Parses XML from a string and retrieves elements or specific attribute values based on the search criteria.

    Args:
        stream (str): XML content as a string.
        search_str (str): XPath expression to find elements.
        namespaces (dict): Namespace prefixes and URIs for XPath queries.
        attrib_type (str, optional): Determines the output format:
            - "" (default): Returns a list of matching ET.Element objects.
            - "text": Returns a list of text content from the matched elements.
            - "value": Returns a list of dictionaries with attributes of the matched elements.

    Returns:
        None | list[ET.Element] | list[dict] | list[str]:
            Depending on `attrib_type`, returns:
            - None if no elements found.
            - List of ET.Element objects if `attrib_type` is empty.
            - List of text strings if `attrib_type` is "text".
            - List of dictionaries of element attributes if `attrib_type` is "value".
    """


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


def gather_namespaces(file: BytesIO) -> dict[str, str]:
    """Extracts XML namespaces from the given file-like object.

    Args:
        file: (BytesIO): XML data in byte format.

    Returns:
        dict[str, str]: A dictionary mapping namespace prefixes to their URIs.
                        The default namespace is mapped to the key 'default'.
    """

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
    """Reads and returns the content of a file inside a ZIP archive as a UTF-8 decoded string.

    Args:
        comp (ZipFile): An open ZipFile object.
        filename (str): The path of the file inside the ZIP archive to read.

    Returns:
        str: The decoded content of the specified file.
    """

    with comp.open(filename) as file:

        return file.read().decode('utf-8')
    

def remove_child_nodes(root: ET.Element, namespaces: dict[str, str], node_name: str) -> ET.Element:
        """Removes every found child node matching `node_name` from the given XML root element.

        Args:
            root (ET.Element): The XML root element from which nodes will be removed.
            namespaces (dict[str, str]): Namespace mappings used for XPath queries.
            node_name (str): XPath expression to find the nodes to remove.

        Returns:
            ET.Element: The modified root element after removing the matching child nodes.
        """

        node = root.findall(node_name, namespaces)[0]
        for child in list(node):
            node.remove(child)
        return root

