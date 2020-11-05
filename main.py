import argparse
import json
import pandas as pd
from glob import glob
from os import path
from more_itertools import flatten

import execjs

import const

# global constant variables
const.vba_module_ext = 'bas'
const.vba_class_ext = 'cls'
const.ignore_filelist_name = 'ignore_file_list.txt'
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

# The vba script order must be module -> property
def change_json_format(json_data: dict, vba_class_name: str) -> dict:
    if len(json_data['Module']) == 0:
        raise ValueError('No Module/Property discription')

    module_list = json_data['Module'].get('Module', [])
    property_list  = json_data['Module'].get('Property', [])

    total_pairs = [('Module', x) for x in module_list] + [('Property', x) for x in property_list]

    procedure_list = json_data['Procedures']
    output = []

    num_of_items = len(total_pairs)

    if num_of_items < len(procedure_list):
        raise ValueError('No Module discription is not enough')
    
    # Convert to markdown format
    output += create_title_output(vba_class_name) # Add title contents
    for i in range(num_of_items):
        output += create_one_module_output(total_pairs[i], procedure_list[i], i) # Add each procedures content
    return json.dumps(output)

def create_title_output(title_name: str) -> list:
    return [{'h1': title_name}]

def create_one_module_output(module_pair: tuple, procedure_dict: dict, step_idx: int) -> list:
    tmp_output = []
    tmp_output.append({'h2' : module_pair[0] + ': ' + procedure_dict['Name']}) # todo: change to acsept property
    tmp_output.append({'p': module_pair[1]})

    tmp_output.append({'h3': '**Basic infomation**'})
    table_dict = set_default_table_dict(['key', 'value'])
    df_tmp = pd.DataFrame([(k,v) for k,v in procedure_dict.items() if k in const.basic_info_cols])
    df_tmp = df_tmp.astype(str)
    table_dict['table']['rows'] = df_tmp.values.tolist()
    tmp_output.append(table_dict)

    if not procedure_dict['Param'] is None:
        tmp_output.append({'h3': '**Parameter**'})

        if isinstance(procedure_dict['Param'], list): # Case of several parameter
            df_tmp = pd.DataFrame(procedure_dict['Param'], 
                                index = [i+1 for i in range(len(procedure_dict['Param']))]).reset_index(drop= False)
        elif isinstance(procedure_dict['Param'],dict): # Case of one parameter
            df_tmp = pd.Series(procedure_dict['Param'], name = 'Param1').to_frame().T.reset_index(drop=False)
        else:
            raise ValueError('Unknown json data')

        df_tmp = df_tmp.astype(str)
        table_dict = set_default_table_dict(df_tmp.columns.to_list()) 
        table_dict['table']['rows'] = df_tmp.to_dict('records')
        tmp_output.append(table_dict)

    tmp_output.append({'p':'---'})
    return tmp_output

def set_default_table_dict(header_names: list) -> dict:
    return {'table': {'headers':header_names, 'rows' :[]}}          

def convert_json_to_md(json_data: dict, vba_class_name: str):
    doc_data_json = change_json_format(json_data, vba_class_name)

    with open(const.convert_code_path) as js_file:
        js_text = js_file.read()
    
    # Compile javascript code
    ctx = execjs.compile(js_text)
    result_md = ctx.call('runJsonToMarkdown', doc_data_json)
    return result_md


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='VBA specification document generater') 
    parser.add_argument('input_file_directory', default = '', help='Directry path for input VBA files') 
    parser.add_argument('output_file_directory', default = '', help='Directry path for output markdown files') 

    args = parser.parse_args() 

    vba_file_name_lists = glob(path.join(args.input_file_directory, f'**.{const.vba_module_ext}')) + \
                          glob(path.join(args.input_file_directory, f'**.{const.vba_class_ext}'))

    try:
        with open(const.ignore_filelist_name) as ig_file:
            ig_list = list(ig_file)
        ig_list = tuple(path.join(args.input_file_directory, f) for f in ig_list)
    except Exception as ex:
        raise Exception (ex)

    for fn in vba_file_name_lists:
        if fn in ig_list:
            continue

        try:
            input_file_name = fn.split('\\')[-1]
            result_json = create_json_from_vba(fn)
            result_md = convert_json_to_md(result_json, input_file_name)
        except Exception as ex:
            raise Exception(fn, ex)
        
        output_file_name = 'doc_' + input_file_name.split('.')[0] +'.md'
        ouput_file_path = path.join(args.output_file_directory, output_file_name)

        with open(ouput_file_path, mode='w') as f:
            f.write(result_md)



