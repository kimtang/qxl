
# qxl

kdb+ integration into excel using exceldna and fsharp.

## Introduction

With qxl you can query kdb+ and display the result in excel.

## Installation

Go to the [relase page](https://github.com/kimtang/qxl/releases) and install the add-in in excel. (Home-> More -> Options -> Add-ins -> Go -> Browse).

![excel add-in!](pic/add-in.jpg "excel add-in")

## Usage

Once the add-ins is installed there are several functions available to communicate with the kdb+ process.

|         Func         |        Description        |  Mode |
|----------------------|---------------------------|-------|
| dna_open_connection  | open a connection to kdb+ | sync  |
| dna_aopen_connection | open a connection to kdb+ | async |
| dna_execute          | execute the query         | sync  |
| dna_axecute          | execute the query         | sync  |
| dna_close_connection | close the connection      | sync  |
| dna_open_subscriber  | subscribe to tick+        | async |
| dna_subscribe        | subscribe to tick+        | async |

## Demo

1. Run `demo.q`
```
$ q demo.q / it will point to port 8866 automatically
```

2. Start excel sheet `demo.xlsx`.
![excel screenshot!](pic/excel-screenshot.jpg "excel screenshot")