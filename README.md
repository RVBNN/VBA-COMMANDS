# VBA-COMMANDS
Most frequent used commands in VBA

## GET FIRST & LAST USED ROW IN A WORKSHEET
R1 = Worksheets("Datos").Cells(1, "Q").End(xlDown).Row 'Fin serie
R2 = Worksheets("Datos").Cells(R1, "Q").End(xlUp).Row   'Inicio serie
