# mettler_toledo_balance_to_excel
A python script to read a Mettler balance and pass the value to Excel via https://github.com/xlwings/xlwings
## Usage  
The python file and the Excel file need to be in the same directory.  
In the Excel worksheet, write a macro with:
```Sub Button1_Click()
    RunPython ("import balance_read_mettler; balance_read_mettler.get_mass()")
    ActiveCell.Offset(0, 1).Select
End Sub```

As written the Excel macro writes the mass passed to it by python, then moves one cell to the right ```ActiveCell.Offset(0, 1).Select```  If you want to move down one cell instead then you would use ```(1, 0)```.    
