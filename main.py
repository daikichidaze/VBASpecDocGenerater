import argparse
import subprocess
from glob import glob
from os import path

import const
from jspy import js
from xdocgen import docgen

def create_json_from_vba(code_path: str):
    with open(code_path) as file:
        text = file.read()

    t = docgen.vbaDocGen(text)
    #t = subprocess.run(['node', 'xdocgen/docgen.js', text], shell=True, stdout=subprocess.PIPE, check=True)
    return t

if __name__ == "__main__":
    print('test')
    const.vba_module_ext = 'bas'
    const.vba_class_ext = 'cls'

    parser = argparse.ArgumentParser(description='VBA specification document generater') 
    parser.add_argument('directory', help='Directry path for VBA files') 

    args = parser.parse_args() 

    vba_file_name_lists = glob(path.join(args.directory, f'**.{const.vba_module_ext}')) + glob(path.join(args.directory, f'**.{const.vba_class_ext}'))

    for path in vba_file_name_lists:
        if 'Button' in path:
            t = create_json_from_vba('xdocgen/test.bas')



