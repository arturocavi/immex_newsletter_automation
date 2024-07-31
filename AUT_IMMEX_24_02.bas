Attribute VB_Name = "AUT_IMMEX_24_02"
' Variables globales
Global Grafica_man
Global Grafica_nom
Global Grafica_acu
Global Grafica_tot_est
Global Grafica_dis_tot_est
Global Grafica_tot_tra
Global Grafica_var_tot_tra
Global Grafica_dist_tot_tra

Sub Macro_inicio()
Attribute Macro_inicio.VB_ProcData.VB_Invoke_Func = "a\n14"

Call WorksheetLoop
Call Generación_Word

End Sub

Sub WorksheetLoop()

        Dim WS_Count As Integer
        Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
        WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
        For I = 1 To WS_Count
            
            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            
            'MsgBox ActiveWorkbook.Worksheets(I).Name
            
            Worksheets(I).Select
            nombre = Worksheets(I).Name
            

            If InStr(1, nombre, "DIST", vbTextCompare) <> 0 Then
                Call Macro_Graficas
            ElseIf InStr(1, nombre, "ACU", vbTextCompare) <> 0 Then
                Call Macro_doble_linea_variacion
            ElseIf InStr(1, nombre, "MEN", vbTextCompare) <> 0 Or InStr(1, nombre, "TOT", vbTextCompare) <> 0 Then
                Call Macro_linea_promedio
            End If
            
        Next I

End Sub


Sub Macro_Graficas()
'
'DIST TOT EST & DIST TOT TRA
'Macro que convierte datos en una gráfica de barras
' ctrl + g

'Dim Grafica_dist_tot_est As ChartObject
'Dim Grafica_dist_tot_tra As ChartObject

If InStr(1, ActiveSheet.Name, "DIST TOT EST", vbBinaryCompare) = 1 Then
    fila = 4
    
    Do While Cells(fila, 1) = ""
        fila = fila + 1
    Loop
    
    ultimo = fila
    Do While Cells(ultimo, 1) <> ""
        ultimo = ultimo + 1
    Loop
    '
    
    '
    Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
    '
    
    Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending
    
    '
    jalisco = fila
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - fila
    '
    Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
    '
    Set Grafica_dis_tot_est = ActiveSheet.ChartObjects.Add(Left:=192, Width:=468.1, Top:=60, Height:=396.9)
    
    With Grafica_dis_tot_est.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
    '
    Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
    '
    Rg = "B" & (fila + 1) & ":B" & (ultimo - 1)
    Range(Rg).NumberFormat = "0.0"

ElseIf InStr(1, ActiveSheet.Name, "DIST TOT TRA", vbBinaryCompare) = 1 Then
    fila = 4
    
    Do While Cells(fila, 1) = ""
        fila = fila + 1
    Loop
    
    ultimo = fila
    Do While Cells(ultimo, 1) <> ""
        ultimo = ultimo + 1
    Loop
    '
    
    '
    Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
    '
    
    Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending
    
    '
    jalisco = fila
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - fila
    '
    Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
    '
    Set Grafica_dist_tot_tra = ActiveSheet.ChartObjects.Add(Left:=192, Width:=468.1, Top:=60, Height:=396.9)
    
    With Grafica_dist_tot_tra.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
    '
    Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
    '
    Rg = "B" & (fila + 1) & ":B" & (ultimo - 1)
    Range(Rg).NumberFormat = "0.0"
End If

End Sub


Sub Macro_linea_promedio()

'Macro para gráficas de barras con linea de promedio
'Ctrl + p

'Dim Grafica_man As ChartObject
'Dim Grafica_nom As ChartObject
'Dim Grafica_tot_est As ChartObject
'Dim Grafica_tot_tra As ChartObject
'Dim Grafica_var_tot_tra As ChartObject


inicio = 5
fin = inicio

Do While Cells(fin, 2) <> ""
    fin = fin + 1
Loop



If InStr(1, ActiveSheet.Name, "MAN EXP MEN", vbBinaryCompare) = 1 Then
    Rg = "C" & (inicio + 1) & ":D" & (fin - 1)
    Range(Rg).NumberFormat = "#,##0"
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_man = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)
    
    With Grafica_man.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT PROM DOR.crtx")
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
    
    Cells(inicio + 1, "G").NumberFormat = "#,##0"
    Cells(inicio + 1, "H").NumberFormat = "#,##0"
    Cells(inicio + 1, "I").NumberFormat = "0.0"
    Cells(inicio + 1, "J").NumberFormat = "0.0"
    Cells(inicio + 1, "K").NumberFormat = "#,##0"
    Cells(inicio + 1, "L").NumberFormat = "#,##0"
ElseIf InStr(1, ActiveSheet.Name, "NOM EXP MEN", vbBinaryCompare) = 1 Then
    Rg = "C" & (inicio + 1) & ":D" & (fin - 1)
    Range(Rg).NumberFormat = "#,##0"
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_nom = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)
    
    With Grafica_nom.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT PROM DOR.crtx")
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
    
    Cells(inicio + 1, "G").NumberFormat = "#,##0"
    Cells(inicio + 1, "H").NumberFormat = "#,##0"
    Cells(inicio + 1, "I").NumberFormat = "0.0"
    Cells(inicio + 1, "J").NumberFormat = "0.0"
    Cells(inicio + 1, "K").NumberFormat = "#,##0"
    Cells(inicio + 1, "L").NumberFormat = "#,##0"
ElseIf InStr(1, ActiveSheet.Name, "TOT EST", vbBinaryCompare) = 1 Then
    Rg = "C" & (inicio + 1) & ":D" & (fin - 1)
    Range(Rg).NumberFormat = "#,##0"
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_tot_est = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)
    
    With Grafica_tot_est.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT PROM DOR.crtx")
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With

    Cells(inicio + 1, "G").NumberFormat = "#,##0"
    Cells(inicio + 1, "H").NumberFormat = "#,##0"
    Cells(inicio + 1, "I").NumberFormat = "#,##0"
    Cells(inicio + 1, "J").NumberFormat = "0.0"
    Cells(inicio + 1, "K").NumberFormat = "#,##0"
    Cells(inicio + 1, "L").NumberFormat = "#,##0"
ElseIf InStr(1, ActiveSheet.Name, "TOT TRA", vbBinaryCompare) = 1 Then
    Rg = "C" & (inicio + 1) & ":D" & (fin - 1)
    Range(Rg).NumberFormat = "#,##0"
          
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_tot_tra = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)
    
    With Grafica_tot_tra.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT PROM DOR.crtx")
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With

      
    Cells(inicio + 1, "G").NumberFormat = "#,##0"
    Cells(inicio + 1, "H").NumberFormat = "#,##0"
    Cells(inicio + 1, "I").NumberFormat = "#,##0"
    Cells(inicio + 1, "J").NumberFormat = "0.0"
    Cells(inicio + 1, "K").NumberFormat = "0.0"
    Cells(inicio + 1, "L").NumberFormat = "#,##0"
    Cells(inicio + 1, "M").NumberFormat = "#,##0"
ElseIf InStr(1, ActiveSheet.Name, "VAR TOT TRA", vbBinaryCompare) = 1 Then
    Rg = "C" & (inicio + 1) & ":D" & (fin - 1)
    Range(Rg).NumberFormat = "0.0"
    Rg2 = "G" & (inicio + 1) & ":K" & (inicio + 1)
    Range(Rg2).NumberFormat = "0.0"
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_var_tot_tra = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)
    
    With Grafica_var_tot_tra.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT PROM DOR.crtx")
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
    
End If

End Sub

Sub Macro_doble_linea_variacion()

'Macro para gráficas de barras con dos lineas de variación, siendo una de ellas promedio
'Ctrl + d

'Dim Grafica_acu As ChartObject

inicio = 5
fin = inicio

Do While Cells(fin, 2) <> ""
    fin = fin + 1
Loop

Range("A" & (inicio) & ":E" & (fin - 1)).Select

Set Grafica_acu = ActiveSheet.ChartObjects.Add(Left:=287, Width:=468.1, Top:=105, Height:=255.1)

With Grafica_acu.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT IMMEX VAR VARPROM.crtx")
    .SetSourceData Source:=Range("A" & (inicio) & ":E" & (fin - 1))
    For k = 1 To (fin - 1)
        If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
            .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
        End If
    Next k
    .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
End With

If InStr(1, ActiveSheet.Name, "ACU", vbTextCompare) <> 0 Then
    Rg = "C" & (inicio + 1) & ":C" & (fin - 1)
    Range(Rg).NumberFormat = "#,##0"
    Cells(inicio + 1, "H").NumberFormat = "#,##0"
    Cells(inicio + 1, "I").NumberFormat = "#,##0"
    Rg2 = "D" & (inicio + 1) & ":E" & (fin - 1)
    Range(Rg2).NumberFormat = "0.0"
    Cells(inicio + 1, "J").NumberFormat = "0.0"
    Cells(inicio + 1, "K").NumberFormat = "0.0"
    Cells(inicio + 1, "L").NumberFormat = "0.0"
End If

End Sub


Sub Generación_Word()

' Nombre  y ubicación de la plantilla
plantilla = "C:\Users\arturo.carrillo\Documents\IMMEX\AUT\PLANTILLA.dotx" ' UBICACIÓN PERSONAL

' Creamos el nuevo archivo word usando la plantilla
Set aplicacion = CreateObject("Word.Application")
aplicacion.Visible = True

Set documento = aplicacion.Documents.Add(Template:=plantilla, NewTemplate:=False, DocumentType:=0)

' Cambiamos la fecha del encabezado
diahoy = Format(Day(Now), "00")
meshoy = Format(Month(Now), "00")
añohoy = Year(Now)
If Month(Now) = 1 Then
    meshoypal = "enero"
    mesbas = Format(11, "00")
    mesbaspal = "noviembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 2 Then
    meshoypal = "febrero"
    mesbas = Format(12, "00")
    mesbaspal = "diciembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 3 Then
    meshoypal = "marzo"
    mesbas = Format(1, "00")
    mesbaspal = "enero"
    añobas = Year(Now)
ElseIf Month(Now) = 4 Then
    meshoypal = "abril"
    mesbas = Format(2, "00")
    mesbaspal = "febrero"
    añobas = Year(Now)
ElseIf Month(Now) = 5 Then
    meshoypal = "mayo"
    mesbas = Format(3, "00")
    mesbaspal = "marzo"
    añobas = Year(Now)
ElseIf Month(Now) = 6 Then
    meshoypal = "junio"
    mesbas = Format(4, "00")
    mesbaspal = "abril"
    añobas = Year(Now)
ElseIf Month(Now) = 7 Then
    meshoypal = "julio"
    mesbas = Format(5, "00")
    mesbaspal = "mayo"
    añobas = Year(Now)
ElseIf Month(Now) = 8 Then
    meshoypal = "agosto"
    mesbas = Format(6, "00")
    mesbaspal = "junio"
    añobas = Year(Now)
ElseIf Month(Now) = 9 Then
    meshoypal = "septiembre"
    mesbas = Format(7, "00")
    mesbaspal = "julio"
    añobas = Year(Now)
ElseIf Month(Now) = 10 Then
    meshoypal = "octubre"
    mesbas = Format(8, "00")
    mesbaspal = "agosto"
    añobas = Year(Now)
ElseIf Month(Now) = 11 Then
    meshoypal = "noviembre"
    mesbas = Format(9, "00")
    mesbaspal = "septiembre"
    añobas = Year(Now)
ElseIf Month(Now) = 12 Then
    meshoypal = "diciembre"
    mesbas = Format(10, "00")
    mesbaspal = "octubre"
    añobas = Year(Now)
End If

'FECHAS MANUALES
'diahoy = InputBox("Ingresa el día de hoy en formato de número a dos dígitos (ej. 23):")'
'meshoy = InputBox("Ingresa el mes de hoy en formato de número a dos dígitos (ej. 10):")
'añohoy = InputBox("Ingresa el año de hoy en formato de número a cuatro dígitos (ej. 2019):")
'meshoypal = InputBox("Ingresa el mes de hoy en formato de palabra en minúsculas (ej. octubre):")
'mesbas = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de número a dos dígitos (ej. 08):")
'mesbaspal = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de palabra en minúsculas (ej. agosto):")
'añobas = InputBox("Ingresa el año de la última base de datos del INEGI (dos meses atrás) en formato de número a cuatro dígitos (ej. 2019):")


' Cambiamos los espaciados del boletín
With documento.Content
    .Style = "Espaciado principal"
End With

' Insertar título del boletín
documento.Content.insertparagraphafter

With documento.Content
    .InsertAfter Hoja9.Cells(2, 1).Value ' Título del boletín [Paragraphs(2)]
    .insertparagraphafter
End With

With documento.Paragraphs(2).Range
    .Style = "Título 1"
End With


' Insertar párrafo de texto1 MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(5, 2).Value ' Texto1 MAN EXP MEN [Paragraphs(4)]
    .insertparagraphafter
End With

With documento.Paragraphs(3).Range
    .Style = "Normal"
End With

' Insertar párrafo de texto1_2 MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(6, 2).Value ' Texto1_2 MAN EXP MEN [Paragraphs(5)]
    .insertparagraphafter
End With

With documento.Paragraphs(4).Range
    .Style = "Normal"
End With

' Insertar título de gráfica MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(7, 2).Value ' Título de gráfica MAN EXP MEN [Paragraphs(6)]
    .insertparagraphafter
End With

With documento.Paragraphs(5).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica MAN EXP MEN
Grafica_man.Chart.ChartArea.Copy
documento.Paragraphs(6).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(8)]
    .insertparagraphafter
End With

With documento.Paragraphs(7).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(8).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter

' Insertar párrafo de texto NOM EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(10, 2).Value ' texto NOM EXP MEN [Paragraphs(11)]
    .insertparagraphafter
End With

With documento.Paragraphs(10).Range
    .Style = "Normal"
End With

' Insertar título de gráfica NOM EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(11, 2).Value ' Título de gráfica NOM EXP MEN [Paragraphs(12)]
    .insertparagraphafter
End With

With documento.Paragraphs(11).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica NOM EXP MEN
Grafica_nom.Chart.ChartArea.Copy
documento.Paragraphs(12).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(14)]
    .insertparagraphafter
End With

With documento.Paragraphs(13).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(14).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto TOT EXP ACU
With documento.Content
    .InsertAfter Hoja9.Cells(14, 2).Value ' Texto TOT EXP ACU [Paragraphs(17)]
    .insertparagraphafter
End With

With documento.Paragraphs(16).Range
    .Style = "Normal"
End With

' Insertar título de gráfica TOT EXP ACU
With documento.Content
    .InsertAfter Hoja9.Cells(15, 2).Value ' Título de gráfica TOT EXP ACU [Paragraphs(18)]
    .insertparagraphafter
End With

With documento.Paragraphs(17).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica TOT EXP ACU
Grafica_acu.Chart.ChartArea.Copy
documento.Paragraphs(18).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(20)]
    .insertparagraphafter
End With

With documento.Paragraphs(19).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja9.Cells(16, 2).Value ' Nota [Paragraphs(21)]
    .insertparagraphafter
End With

With documento.Paragraphs(20).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(21).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(19, 2).Value ' Texto TOT EST [Paragraphs(24)]
    .insertparagraphafter
End With

With documento.Paragraphs(23).Range
    .Style = "Normal"
End With

' Insertar título de gráfica TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(20, 2).Value ' Título de gráfica TOT EST [Paragraphs(25)]
    .insertparagraphafter
End With

With documento.Paragraphs(24).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica TOT EST
Grafica_tot_est.Chart.ChartArea.Copy
documento.Paragraphs(25).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(27)]
    .insertparagraphafter
End With

With documento.Paragraphs(26).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja9.Cells(21, 2).Value ' Nota [Paragraphs(28)]
    .insertparagraphafter
End With

With documento.Paragraphs(27).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(28).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto DIST TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(24, 2).Value ' Texto DIST TOT EST [Paragraphs(31)]
    .insertparagraphafter
End With

With documento.Paragraphs(30).Range
    .Style = "Normal"
End With

' Insertar título de gráfica DIST TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(25, 2).Value ' Título de la gráfica DIST TOT EST [Paragraphs(32)]
    .insertparagraphafter
End With

With documento.Paragraphs(31).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica DIST TOT EST
Grafica_dis_tot_est.Chart.ChartArea.Copy
documento.Paragraphs(32).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(34)]
    .insertparagraphafter
End With

With documento.Paragraphs(33).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(34).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto1 TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(28, 2).Value ' Texto1 TOT TRA [Paragraphs(37)]
    .insertparagraphafter
End With

With documento.Paragraphs(36).Range
    .Style = "Normal"
End With

' Insertar texto_2 TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(29, 2).Value ' Texto_2 TOT TRA [Paragraphs(38)]
    .insertparagraphafter
End With

With documento.Paragraphs(37).Range
    .Style = "Normal"
End With


' Insertar título de gráfica TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(30, 2).Value ' Título de la gráfica TOT TRA [Paragraphs(39)]
    .insertparagraphafter
End With

With documento.Paragraphs(38).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica TOT TRA
Grafica_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(39).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(41)]
    .insertparagraphafter
End With

With documento.Paragraphs(40).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja9.Cells(31, 2).Value ' Nota [Paragraphs(42)]
    .insertparagraphafter
End With

With documento.Paragraphs(41).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(42).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter


' Insertar texto VAR TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(34, 2).Value ' Texto VAR TOT TRA [Paragraphs(45)]
    .insertparagraphafter
End With

With documento.Paragraphs(44).Range
    .Style = "Normal"
End With

' Insertar título de gráfica VAR TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(35, 2).Value ' Título de la gráfica VAR TOT TRA [Paragraphs(46)]
    .insertparagraphafter
End With

With documento.Paragraphs(45).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica VAR TOT TRA
Grafica_var_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(46).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(48)]
    .insertparagraphafter
End With

With documento.Paragraphs(47).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(48).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter


' Insertar texto DIST TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(38, 2).Value ' Texto VAR TOT TRA [Paragraphs(51)]
    .insertparagraphafter
End With

With documento.Paragraphs(50).Range
    .Style = "Normal"
End With

' Insertar título de gráfica DIST TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(39, 2).Value ' Título de la gráfica DIST TOT TRA [Paragraphs(52)]
    .insertparagraphafter
End With

With documento.Paragraphs(51).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica DIST TOT TRA
Grafica_dist_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(52).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI." ' Nota [Paragraphs(54)]
    .insertparagraphafter
End With

With documento.Paragraphs(53).Range
    .Style = "Fuentes"
End With


' Cambiar la fecha de realización
Set cuadrofecha = documento.Sections(1).Headers(1).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  350, 48 - 7, 240, 37.5 / 2)
                  ' wdHeaderFooterPrimary = 1
cuadrofecha.TextFrame.TextRange.Text = "Ficha informativa, " & diahoy & " de " & meshoypal & " de " & añohoy
cuadrofecha.TextFrame.TextRange.Font.Color = RGB(98, 113, 120)
cuadrofecha.TextFrame.TextRange.Font.Underline = wdUnderlineSingle
cuadrofecha.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
cuadrofecha.Fill.ForeColor = RGB(255, 255, 255)
cuadrofecha.Line.ForeColor = RGB(255, 255, 255)

' Guardar el documento (escribir la dirección en donde se quiera guardar)
documento.SaveAs "C:\Users\arturo.carrillo\Documents\IMMEX\" & añobas & " " & mesbas & "\Ficha informativa IMMEX " & mesbaspal & " " & añobas & "-" & añohoy & meshoy & diahoy & ".docx" ' UBICACIÓN PERSONAL

End Sub

