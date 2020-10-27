import argparse
import json
import pandas as pd
from glob import glob
from os import path

import execjs

import const

const.vba_module_ext = 'bas'
const.vba_class_ext = 'cls'
const.js_code_path = 'xdocgen/docgen.js'
const.convert_code_path = 'json2md/run_json2md.js'

def create_json_from_vba(vba_code_path: str):
    with open(vba_code_path) as vba_file:
        vba_text = vba_file.read()
    
    with open(const.js_code_path) as js_file:
        js_text = js_file.read()
    
    # Compile javascript code
    ctx = execjs.compile(js_text)
    result_json = ctx.call('vbaDocGen', vba_text)
    return result_json

def convert_json_to_md(json_data):
    output = []
    for k,v in json_data.items():
        if isinstance(v,dict):
            for k2, v2 in v.items():
                output.append({'h1' : k2})
                output.append({'h2': v2})
        elif isinstance(v, list):
            for itm in v:
                output.append({'h2': 'Basic infomation'})
                table_dict = {'table': {'headers':['key', 'value'],
                                        'rows' :[]}}
                for k3,v3 in itm.items():
                    if k3 in 'Name Scope Static Procedure Type Description Retuens'.split():
                        table_dict['table']['rows'].append([k3,str(v3)])
                output.append(table_dict)
                
                for k3,v3 in itm.items():
                    if k3 == 'Param':
                        output.append({'h2': 'Params'})
                        df = pd.DataFrame(v3, index = [f'Param{i+1}' for i in range(len(v3))]).reset_index(drop= False)
                        df = df.astype(str)
                        table_dict = {'table' : {}}
                        table_dict['table']['headers'] = df.columns.to_list()
                        table_dict['table']['rows'] = df.to_dict('records')
                output.append(table_dict)

    json_str = json.dumps(output)
#     return json_str    

    with open(const.convert_code_path) as js_file:
        js_text = js_file.read()
    
    # Compile javascript code
    ctx = execjs.compile(js_text)
    result_md = ctx.call('runJsonToMarkdown', json_str)
    return result_md


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='VBA specification document generater') 
    parser.add_argument('directory', help='Directry path for VBA files') 

    args = parser.parse_args() 

    vba_file_name_lists = glob(path.join(args.directory, f'**.{const.vba_module_ext}')) + glob(path.join(args.directory, f'**.{const.vba_class_ext}'))

    for path in vba_file_name_lists:
        if 'Button' in path:
            result_json_str = create_json_from_vba('xdocgen/test.bas')
            result_md = convert_json_to_md(json.loads(result_json_str))
    
    with open('test.md', mode='w') as f:
        f.write(result_md)
            



