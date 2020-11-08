# SpecDocGenerater
Specefication document generater for VBA code

# Overview
* Python script to generate a specification documor VBA in markdown format
* 2 javascript modules are run via PyExecJS
  * PyExecJS are already end of life (https://pypi.org/project/PyExecJS/)
* Appreciated for the each modules in bellow
  * @x-vba [XDocGen](https://github.com/x-vba/xdocgen)
    * Javascript module to create a json data from VBA source code comment
  * @IonicaBizau [json2md](https://github.com/IonicaBizau/json2md)
    * Javascript module to convert the json data to markdown data

# Description
* Specification document content will be generate from the comments in VBA source code
* The document will be output by markdown format
* Python script will call 2 javascript modules
  * XDocGen
  * json2md
* Please follow the syntax in [XDocGen homepage](https://x-vba.com/xdocgen/) for the format of VBA source code comments

# How to use
## Install
### PyExecJS

```
pip install PyExecJS
```

### json2md

```
yarn install indento
yarn install json2md
```
*Developed in windows 10 environment. Please change the installation command based on your environment*

## Add comments in VBA source code
* Basically follows the syntax in [XDocGen homepage](https://x-vba.com/xdocgen/)
* "@Module" and "@Property" are only allowed in the **Module Level Tags**

### Example
example

## Execute python code

```
python main.py [input folder directory] [output folder directory]
```