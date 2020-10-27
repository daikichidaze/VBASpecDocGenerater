import argparse
import json
import pandas as pd
from glob import glob
from os import path

import execjs

import const

# global constant variables
const.vba_module_ext = 'bas'
const.vba_class_ext = 'cls'
const.docgen_js_code_path = 'xdocgen/docgen.js'
const.docgen_func_name = 'vbaDocGen'
const.convert_code_path = 'json2md/run_json2md.js'
const.basic_info_cols = ['Name', 'Scope', 'Static', 'Procedure', 'Type', 'Description', 'Returns']

# use javascript for vba comment -> json
def create_json_from_vba(vba_code_path: str) -> dict:
    with open(vba_code_path) as vba_file:
        vba_text = vba_file.read()
    
    with open(const.docgen_js_code_path) as js_file:
        js_text = js_file.read()
    
    # Compile javascript code
    ctx = execjs.compile(js_text)
    result_str = ctx.call(const.docgen_func_name, vba_text)
    return json.loads(result_str)

def change_json_format(json_data: dict) -> dict:
    module_list = json_data['Module']['Module']
    procedure_list = json_data['Procedures']
    output = []

    num_of_modules = len(module_list)
    
    for i in range(num_of_modules):
        output += create_one_module_output(module_list[i], procedure_list[i], i)
    return json.dumps(output)

def create_one_module_output(module_str: str, procedure_dict: dict, step_idx: int) -> list:
    tmp_output = []
    tmp_output.append({'h1' : 'Module: ' + procedure_dict['Name']})
    tmp_output.append({'h2': module_str})

    tmp_output.append({'h3': 'Basic infomation'})
    table_dict = set_default_table_dict(['key', 'value'])
    df_tmp = pd.DataFrame([(k,v) for k,v in procedure_dict.items() if k in const.basic_info_cols])
    df_tmp = df_tmp.astype(str)
    table_dict['table']['rows'] = df_tmp.values.tolist()
    tmp_output.append(table_dict)

    if not procedure_dict['Param'] is None:
        tmp_output.append({'h3': 'Parameter'})
        df_tmp = pd.DataFrame(procedure_dict['Param'], 
                            index = [f'Param{i+1}' for i in range(len(procedure_dict['Param']))]).reset_index(drop= False)
        df_tmp = df_tmp.astype(str)
        table_dict = set_default_table_dict(df_tmp.columns.to_list()) 
        table_dict['table']['rows'] = df_tmp.to_dict('records')
        tmp_output.append(table_dict)

    return tmp_output

def set_default_table_dict(header_names: list) -> dict:
    return {'table': {'headers':header_names, 'rows' :[]}}          

def convert_json_to_md(json_data : dict):
    doc_data_json = change_json_format(json_data)


    with open(const.convert_code_path) as js_file:
        js_text = js_file.read()
    
    # Compile javascript code
    ctx = execjs.compile(js_text)
    result_md = ctx.call('runJsonToMarkdown', doc_data_json)
    return result_md


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='VBA specification document generater') 
    parser.add_argument('directory', help='Directry path for VBA files') 

    args = parser.parse_args() 

    vba_file_name_lists = glob(path.join(args.directory, f'**.{const.vba_module_ext}')) + glob(path.join(args.directory, f'**.{const.vba_class_ext}'))

    for path in vba_file_name_lists:
        if 'Button' in path:
            result_json = create_json_from_vba('xdocgen/test.bas')
            result_md = convert_json_to_md(result_json)
    
    with open('test.md', mode='w') as f:
        f.write(result_md)
            



