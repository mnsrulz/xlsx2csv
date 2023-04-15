# xlsx2csv
A dotnet tool to convert xlsx files to csv format. Handles large XLSX files. Fast and easy to use.

[![license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/mnsrulz/xlsx2csv/blob/master/LICENSE)
[![nuget](https://img.shields.io/nuget/v/xlsx2csv.svg)](https://www.nuget.org/packages/xlsx2csv)
[![downloads](https://img.shields.io/nuget/dt/xlsx2csv.svg)](https://www.npmjs.com/package/nurlresolver)
[![github forks](https://img.shields.io/github/forks/mnsrulz/xlsx2csv.svg)](https://github.com/mnsrulz/xlsx2csv/network/members)
[![github stars](https://img.shields.io/github/stars/mnsrulz/xlsx2csv.svg)](https://github.com/mnsrulz/xlsx2csv/stargazers)


# Install
```
dotnet tool install --global xlsx2csv
```

# Usage
```
xlsx2csv file.xlsx 
         [output.csv] 
         [-n sheet1]
         [-d ,] 
         [--help]   
         [--version] 
```
## positional arguments

| Name      | Type | Description     |
| :---        |    :----:   |          ---: 
| xlsxfile      | string       | xlsx file path
| outfile   | string        | (optional) output csv file path


## optional arguments
| Name      | Type | Description     |
| :---        |    :----:   |          ---: 
| -n, --sheetname      | string       | Worksheet name to be processed
| -d, --delimiter   | string        | CSV file separator
| --help   | string        | Display this help screen
| --version   | string        | Display version information