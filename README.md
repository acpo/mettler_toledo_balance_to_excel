# mettler_toledo_balance_to_excel
A python script to read a Mettler balance and pass the value to Excel via https://github.com/xlwings/xlwings  
An open-source version using Libre Office and Python is in [another repository](https://github.com/acpo/mettler_toledo_balance_to_libreoffice).  

## Usage  
The python file and the Excel file need to be in the same directory.  
In the Excel worksheet, assign a keystroke or make a button and assign a macro with the following:
```
Sub Button1_Click()
    RunPython ("import balance_read_mettler; balance_read_mettler.get_mass()")
    ActiveCell.Offset(0, 1).Select
End Sub
```

As written the Excel macro writes the mass passed to it by python, then moves one cell to the right `ActiveCell.Offset(0, 1).Select`  If you want to move down one cell instead then you would use `(1, 0)`.    

## Requirements  
- xlWings addin for Excel (`xlwings.xlam`).  Read the instructions at https://docs.xlwings.org/en/stable/addin.html#xlwings-addin  
- xlwings Python library (usually `pip install xlwings`)  
This was tested on a Mettler Toledo XS105 connected via an RS-232 to USB cable on a Windows 7 system using Python 2.7.  

## Port settings  
The port settings are in the .py file.  You will likely need to change the `port='COM1'` to your relevant COM port.  This is for a Windows OS.  The relevant port settings for MacOS or Linux are commented out in the .py file.  

## Other good ideas  
https://github.com/janelia-pypi/mettler_toledo_device_python  has some good, clever, and better ideas about how to communicate with a Mettler balance.  
