# HRScript
HR Script for my wife.

### What For?
It is used to analyze wife's excel documents, and make her more easier.

1. It will analyze all .xlsx files in `./文件池` folder.
2. Each files has been analyzed, will move into `./文件池/已完成` folder.
3. It will generate a 结果.xls file in `./`, with all informations.

### Dependency
- Python 3+
- xlrd
- xlwt

### Install Dependency
```
$: pip install xlrd
$: pip install xlwt
```

### How To Run
```
$: python main.py
```

### Generate Executable File

#### Dependency
- pyinstaller

#### How To Use

Run `build.bat` script in terminal.
