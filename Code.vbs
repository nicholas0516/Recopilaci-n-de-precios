Sub CopiarArchivosSeleccionados()

 

    Dim rutaDestino As String 'ruta de la carpeta de destino
    Dim archivo As String 'nombre del archivo de origen
    Dim libro As Workbook 'libro de Excel de origen
    Dim hoja As Worksheet 'hoja de Excel de destino
    Dim i As Integer 'contador
    Dim seleccion As Variant 'array de archivos seleccionados

    'definir ruta de destino
    rutaDestino = ThisWorkbook.Path & "\"

    'seleccionar los archivos de Excel de origen
    seleccion = Application.GetOpenFilename(FileFilter:="Archivos de Excel, *.xls*;*.xlsm", Title:="Selecciona los archivos que deseas copiar", MultiSelect:=True)

    'loop para copiar todos los archivos de Excel seleccionados
    For i = LBound(seleccion) To UBound(seleccion)
        'abrir el archivo de Excel de origen
        Set libro = Workbooks.Open(seleccion(i), UpdateLinks:=False)

        'copiar la información a una nueva hoja de Excel en el libro de destino
        Set hoja = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        hoja.Name = Left(libro.Name, Len(libro.Name) - 4) 'asignar el nombre de la hoja según el nombre del archivo de origen
        libro.Sheets(2).Cells.Copy hoja.Range("A1")

        'cerrar el libro de Excel de origen sin guardar los cambios
        libro.Close SaveChanges:=False
    Next i

    'mensaje de finalización
    MsgBox "La copia de archivos de Excel ha finalizado."

End Sub

 

Sub CopiarInformacionHorizontal2()
    Dim datosSheet As Worksheet
    Dim origenSheet As Worksheet
    Dim columnaOrigen As Range
    Dim columnaDestino As Range
    Dim columnaDestino2 As Range
    Dim columnaDestino3 As Range
    Dim filaTitulo As Long
    Dim i As Integer
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim primeraHoja As Boolean
    Dim lastRow As Long
    Dim codigoColumna As Range

    ' Establecer la hoja de destino
    Set datosSheet = ThisWorkbook.Sheets("Datos")

    ' Establecer la columna de destino inicial para los títulos
    Set columnaDestino = datosSheet.Range("A2")

    ' Establecer la fila de título inicial
    filaTitulo = 2

    k = 3
    j = 7
    l = 7

    ' Bandera para verificar si es la primera hoja
    primeraHoja = True

    ' Iterar por todas las hojas del libro
    For Each origenSheet In ThisWorkbook.Sheets
        ' Verificar si el nombre de la hoja cumple con ciertos criterios
        If Left(origenSheet.Name, 4) = "Ene_" Or origenSheet.Name = "Dic2022" Or Right(origenSheet.Name, 4) = "2023" Then

            Set columnaDestino1 = datosSheet.Cells(2, j)

            ' Copiar el nombre de la hoja de origen en la fila anterior de la hoja "Datos"
            datosSheet.Cells(filaTitulo - 1, columnaDestino1.Column).Value = origenSheet.Name
            datosSheet.Cells(filaTitulo - 1, columnaDestino1.Column).Resize(1, 25).Merge


            If primeraHoja Then
                ' Buscar y copiar los títulos de las columnas principales de la primera hoja
                Set columnaOrigen = origenSheet.Rows(1).Find("Marca")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino

                Set columnaOrigen = origenSheet.Rows(1).Find("Clase")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino.Offset(0, 1)

                Set columnaOrigen = origenSheet.Rows(1).Find("Codigo")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino.Offset(0, 2)

                Set columnaOrigen = origenSheet.Rows(1).Find("Referencia1")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino.Offset(0, 3)

                Set columnaOrigen = origenSheet.Rows(1).Find("Referencia2")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino.Offset(0, 4)

                Set columnaOrigen = origenSheet.Rows(1).Find("Referencia3")
                origenSheet.Range(columnaOrigen, columnaOrigen.End(xlDown)).Copy columnaDestino.Offset(0, 5)

                primeraHoja = False
            End If



            ' Copiar y pegar los títulos de los años en la fila 2
            For i = 2000 To 2024
                Set columnaDestino3 = datosSheet.Cells(2, l)
                columnaDestino3.Value = CStr(i)
                l = l + 1
            Next i

            ' Iterar por los códigos en la hoja "Datos"
            Dim codigo As Range
            Set codigo = datosSheet.Range("C3:C" & datosSheet.Cells(datosSheet.Rows.Count, "C").End(xlUp).Row)
            Set codigoColumna = datosSheet.Columns("C")
            lastRow = codigoColumna.Cells(codigoColumna.Rows.Count, 1).End(xlUp).Row

            For Each codigo In codigo
                Dim codigoActual As String
                codigoActual = codigo.Value

                ' Buscar el código en las hojas a partir de la cuarta hoja
                If origenSheet.Index > 3 Then
                    Set columnaDestino2 = datosSheet.Cells(k, j)
                    ' Buscar el código en la columna "Codigo" de la hoja de origen
                    Set columnaOrigen = origenSheet.Rows(1).Find("Codigo")
                    Set filaCodigo = origenSheet.Columns(columnaOrigen.Column).Find(What:=codigoActual, LookIn:=xlValues, LookAt:=xlWhole)

                    If Not filaCodigo Is Nothing Then
                        ' Copiar y pegar la información de los años 2000 a 2024 en la fila del código
                        For i = 2000 To 2024
                            Set columnaOrigen = origenSheet.Rows(1).Find(CStr(i))
                            If Not columnaOrigen Is Nothing Then
                                origenSheet.Cells(filaCodigo.Row, columnaOrigen.Column).Copy columnaDestino2
                                Set columnaDestino2 = columnaDestino2.Offset(0, 1)
                            End If
                        Next i
                        k = k + 1
                    End If
                End If
            Next codigo

            ' Mover la columna de destino al siguiente bloque de años
            j = j + 25
            k = 3
        End If
    Next origenSheet

    ' Ajustar el ancho de las columnas copiadas en la hoja "Datos"
    datosSheet.Columns.AutoFit

    ' Mostrar mensaje de finalización
    MsgBox "La información se ha copiado correctamente.", vbInformation
End Sub

 


Sub CalcDepreciacion()
    Dim lastRow As Long
    Dim wsDatos As Worksheet
    Dim wsDepreciacion As Worksheet
    Dim avgRow As Long
    Dim rng As Range
    Dim cell As Range

 

 

    Set wsDatos = Sheets("Datos")
    Set wsDepreciacion = Sheets("Depreciacion")

 

 

    ' Copiar datos y encabezados
    wsDatos.Columns("A:F").Copy Destination:=wsDepreciacion.Columns("A:A")
    wsDatos.Rows("1:2").Copy Destination:=wsDepreciacion.Range("A1")

 

 

    ' Obtener la última fila con datos en la columna A de la hoja Datos
    lastRow = wsDatos.Cells(wsDatos.Rows.Count, "A").End(xlUp).Row

 

 

    ' Llenar las fórmulas hacia abajo en la hoja Depreciacion
    wsDepreciacion.Range("AF3:BD" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-25]-1, ""0%"")"
    wsDepreciacion.Range("BE3:CC" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-50]-1, ""0%"")"
    wsDepreciacion.Range("CD3:DB" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-75]-1, ""0%"")"
    wsDepreciacion.Range("DC3:EA" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-100]-1, ""0%"")"
    wsDepreciacion.Range("EB3:EZ" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-125]-1, ""0%"")"
    wsDepreciacion.Range("FA3:FY" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-150]-1, ""0%"")"
    wsDepreciacion.Range("FZ3:GX" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-175]-1, ""0%"")"
    wsDepreciacion.Range("GY3:HW" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-200]-1, ""0%"")"
    wsDepreciacion.Range("HX3:IV" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-225]-1, ""0%"")"
    wsDepreciacion.Range("IW3:JU" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-250]-1, ""0%"")"
    wsDepreciacion.Range("JV3:KT" & lastRow).FormulaR1C1 = "=IFERROR(Datos!RC/Datos!RC[-275]-1, ""0%"")"

 

    ' Calcular promedios en filas adicionales
    avgRow = lastRow + 2
    wsDepreciacion.Range("AF" & avgRow & ":BD" & avgRow).Formula = "=AVERAGE(AF3:AF" & lastRow & ")"
    wsDepreciacion.Range("BE" & avgRow & ":CC" & avgRow).Formula = "=AVERAGE(BE3:BE" & lastRow & ")"
    wsDepreciacion.Range("CD" & avgRow & ":DB" & avgRow).Formula = "=AVERAGE(CD3:CD" & lastRow & ")"
    wsDepreciacion.Range("DC" & avgRow & ":EA" & avgRow).Formula = "=AVERAGE(DC3:DC" & lastRow & ")"
    wsDepreciacion.Range("EB" & avgRow & ":EZ" & avgRow).Formula = "=AVERAGE(EB3:EB" & lastRow & ")"
    wsDepreciacion.Range("FA" & avgRow & ":FY" & avgRow).Formula = "=AVERAGE(FA3:FA" & lastRow & ")"
    wsDepreciacion.Range("FZ" & avgRow & ":GX" & avgRow).Formula = "=AVERAGE(FZ3:FZ" & lastRow & ")"
    wsDepreciacion.Range("GY" & avgRow & ":HW" & avgRow).Formula = "=AVERAGE(GY3:GY" & lastRow & ")"
    wsDepreciacion.Range("HX" & avgRow & ":IV" & avgRow).Formula = "=AVERAGE(HX3:HX" & lastRow & ")"
    wsDepreciacion.Range("IW" & avgRow & ":JU" & avgRow).Formula = "=AVERAGE(IW3:IW" & lastRow & ")"
    wsDepreciacion.Range("JV" & avgRow & ":KT" & avgRow).Formula = "=AVERAGE(JV3:JV" & lastRow & ")"

 

 

    ' Autoajustar las columnas
    wsDepreciacion.Columns.AutoFit

 

 

    ' Merge de celdas y agregar el texto "Promedio"
    With wsDepreciacion.Range("A" & avgRow & ":F" & avgRow)
        .Merge
        .Value = "Promedio"
        .HorizontalAlignment = xlCenter
    End With

End Sub
