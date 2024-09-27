from ctypes import Array
import json
from pprint import pprint
from typing import Optional, Union, Literal, Dict, List, Any
import pandas as pd


class Excel_file:
    def __init__(self, excel_file: str):
        self.excel_file = excel_file
        self.excel_frame = pd.read_excel(self.excel_file)
        self.size = self.excel_frame.shape[0]
        self.columns = self.excel_frame.columns
        self.excel_dict = self.excel_frame.to_dict(orient="records")

    def read_excel(self):
        # read excel columns first
        with open("ok.json", "wb") as f:
            f.write(json.dumps(self.excel_dict, indent=4).encode("utf-8"))

        return self.excel_dict

    def create_index(self, create_json: Optional[bool] = None) -> Dict[int, List[Dict]]:
        """
        Creates an index from the Excel data based on section values and optionally writes it to a JSON file.

        Parameters:
        create_json (Optional[bool]): Flag to create a JSON output file if set to True. Default is None.

        Returns:
        Dict[int, List[Dict]]: A dictionary where:
            - Keys are integers derived from the section values.
            - Values are lists containing the row dictionaries from the Excel data corresponding to each section.

        Example:
            0: [
                {
                    \n'Tipo': str,
                    \n'Doc / BOM': int,
                    \n'Descrição': str,
                    \n'Rev.': int,
                    \n'Nome do Arquivo PDF': str,
                    \n'Seção': str ('I-0' | other single letter),
                    \n'Status': str,
                    \n'Obs': str,
                    \n'Unnamed: 8': NaN,
                    \n'Unnamed: 9': NaN
                }
            ]
        """
        mapped_index: Dict[int, List[Dict]] = {}
        for index, item in enumerate(self.excel_dict):
            not_cases = ["NaN", None]
            section = item["Seção"]

            if (
                section is not None
                and section not in not_cases
                and not isinstance(section, (int, float))
            ):
                # creates the summary

                if "I" in section:
                    lookingfor = section.split("-")[-1]
                    try:
                        current_i = int(lookingfor)
                    except ValueError:
                        current_i = None
                    # if the section is like 'I1-I2' or 'I1-I3' ...

                    if current_i is not None:
                        if current_i not in mapped_index:
                            mapped_index[current_i] = []
                        mapped_index[current_i].append(item)

            # TODO
            # get the files after the section definitions
            # get the section files
            # getting an error trying to identify if it's int or not
            if isinstance(section, (int, float)):
                print(f"section {section}")
                print(f"item {item["Descrição"]}")
        pprint(mapped_index)
        if create_json:
            with open("index.json", "wb") as f:
                f.write(json.dumps(mapped_index, indent=4).encode("utf-8"))

        return mapped_index
    
    
    def make_index_page(self):
        mapped_index: Dict[int, List[Dict]] = {}
        cases = ['S', 's']
        all_items: List[Dict[str, Any]] = []
        for index, item in enumerate(self.excel_dict):
            sect = item["Seção"]
            name = item['Nome do Arquivo PDF']
            desc = item["Descrição"]

            if sect in cases or isinstance(sect, int):
                newobj= {
                    "name": name,  
                    "description": desc,
                    "section": sect,
                    "children": []
                }               
                all_items.append(newobj)
            
            
        parents = [item for item in all_items if item['section'] == 'S']
        sub_parents = [item for item in all_items if item['section'] == 's']
        children = [item for item in all_items if isinstance(item['section'], int)]

        # print('INDEX FROM EXCEL')
        for i in children:
            for j in sub_parents:
                if "SEÇÃO " + str(i['section']) == str(j['name']).split('.')[0]:
                    j['children'].append(i)
        
        for i in sub_parents:
            for j in parents:
                if str(j['name']).split('.')[0] == str(i['name']).split('.')[0]:
                    j['children'].append(i)
        pprint(parents)
        return all_items
