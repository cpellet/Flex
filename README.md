# Flex

Foobar is data-science oriented spreadsheet desktop program written in python, with a TkInter GUI.

## Installation

Download a [release](https://github.com/cpellet/Flex/releases) to install Flex and install the following dependencies

```bash
pip install xlsxwriter tksheet
```
Run the script using python 3:
```bash
python3 flex.py
```

## Usage

Flex supports formula entry using the `=` prefix in any cell. The additional features are now supported:
* Data import from other spreadsheet programs such as MS Excel or in csv format
* Support for all functions implemented in python's `math` module
* Cross-cell referencing by address (e.g: `B14`)
* Automatic update propagation to other referenced cells

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://choosealicense.com/licenses/mit/)
