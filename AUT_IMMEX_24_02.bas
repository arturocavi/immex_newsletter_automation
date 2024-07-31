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
Call Generaci�n_Word

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
'Macro que convierte datos en una gr�fica de barras
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

'Macro para gr�ficas de barras con linea de promedio
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

'Macro para gr�ficas de barras con dos lineas de variaci�n, siendo una de ellas promedio
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


Sub Generaci�n_Word()

' Nombre  y ubicaci�n de la plantilla
plantilla = "C:\Users\arturo.carrillo\Documents\IMMEX\AUT\PLANTILLA.dotx" ' UBICACI�N PERSONAL

' Creamos el nuevo archivo word usando la plantilla
Set aplicacion = CreateObject("Word.Application")
aplicacion.Visible = True

Set documento = aplicacion.Documents.Add(Template:=plantilla, NewTemplate:=False, DocumentType:=0)

' Cambiamos la fecha del encabezado
diahoy = Format(Day(Now), "00")
meshoy = Format(Month(Now), "00")
a�ohoy = Year(Now)
If Month(Now) = 1 Then
    meshoypal = "enero"
    mesbas = Format(11, "00")
    mesbaspal = "noviembre"
    a�obas = Year(Now) - 1
ElseIf Month(Now) = 2 Then
    meshoypal = "febrero"
    mesbas = Format(12, "00")
    mesbaspal = "diciembre"
    a�obas = Year(Now) - 1
ElseIf Month(Now) = 3 Then
    meshoypal = "marzo"
    mesbas = Format(1, "00")
    mesbaspal = "enero"
    a�obas = Year(Now)
ElseIf Month(Now) = 4 Then
    meshoypal = "abril"
    mesbas = Format(2, "00")
    mesbaspal = "febrero"
    a�obas = Year(Now)
ElseIf Month(Now) = 5 Then
    meshoypal = "mayo"
    mesbas = Format(3, "00")
    mesbaspal = "marzo"
    a�obas = Year(Now)
ElseIf Month(Now) = 6 Then
    meshoypal = "junio"
    mesbas = Format(4, "00")
    mesbaspal = "abril"
    a�obas = Year(Now)
ElseIf Month(Now) = 7 Then
    meshoypal = "julio"
    mesbas = Format(5, "00")
    mesbaspal = "mayo"
    a�obas = Year(Now)
ElseIf Month(Now) = 8 Then
    meshoypal = "agosto"
    mesbas = Format(6, "00")
    mesbaspal = "junio"
    a�obas = Year(Now)
ElseIf Month(Now) = 9 Then
    meshoypal = "septiembre"
    mesbas = Format(7, "00")
    mesbaspal = "julio"
    a�obas = Year(Now)
ElseIf Month(Now) = 10 Then
    meshoypal = "octubre"
    mesbas = Format(8, "00")
    mesbaspal = "agosto"
    a�obas = Year(Now)
ElseIf Month(Now) = 11 Then
    meshoypal = "noviembre"
    mesbas = Format(9, "00")
    mesbaspal = "septiembre"
    a�obas = Year(Now)
ElseIf Month(Now) = 12 Then
    meshoypal = "diciembre"
    mesbas = Format(10, "00")
    mesbaspal = "octubre"
    a�obas = Year(Now)
End If

'FECHAS MANUALES
'diahoy = InputBox("Ingresa el d�a de hoy en formato de n�mero a dos d�gitos (ej. 23):")'
'meshoy = InputBox("Ingresa el mes de hoy en formato de n�mero a dos d�gitos (ej. 10):")
'a�ohoy = InputBox("Ingresa el a�o de hoy en formato de n�mero a cuatro d�gitos (ej. 2019):")
'meshoypal = InputBox("Ingresa el mes de hoy en formato de palabra en min�sculas (ej. octubre):")
'mesbas = InputBox("Ingresa el mes de la �ltima base de datos del INEGI (dos meses atr�s) en formato de n�mero a dos d�gitos (ej. 08):")
'mesbaspal = InputBox("Ingresa el mes de la �ltima base de datos del INEGI (dos meses atr�s) en formato de palabra en min�sculas (ej. agosto):")
'a�obas = InputBox("Ingresa el a�o de la �ltima base de datos del INEGI (dos meses atr�s) en formato de n�mero a cuatro d�gitos (ej. 2019):")


' Cambiamos los espaciados del bolet�n
With documento.Content
    .Style = "Espaciado principal"
End With

' Insertar t�tulo del bolet�n
documento.Content.insertparagraphafter

With documento.Content
    .InsertAfter Hoja9.Cells(2, 1).Value ' T�tulo del bolet�n [Paragraphs(2)]
    .insertparagraphafter
End With

With documento.Paragraphs(2).Range
    .Style = "T�tulo 1"
End With


' Insertar p�rrafo de texto1 MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(5, 2).Value ' Texto1 MAN EXP MEN [Paragraphs(4)]
    .insertparagraphafter
End With

With documento.Paragraphs(3).Range
    .Style = "Normal"
End With

' Insertar p�rrafo de texto1_2 MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(6, 2).Value ' Texto1_2 MAN EXP MEN [Paragraphs(5)]
    .insertparagraphafter
End With

With documento.Paragraphs(4).Range
    .Style = "Normal"
End With

' Insertar t�tulo de gr�fica MAN EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(7, 2).Value ' T�tulo de gr�fica MAN EXP MEN [Paragraphs(6)]
    .insertparagraphafter
End With

With documento.Paragraphs(5).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica MAN EXP MEN
Grafica_man.Chart.ChartArea.Copy
documento.Paragraphs(6).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(8)]
    .insertparagraphafter
End With

With documento.Paragraphs(7).Range
    .Style = "Fuentes"
End With
' Insertar salto de p�gina
documento.Paragraphs(8).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter

' Insertar p�rrafo de texto NOM EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(10, 2).Value ' texto NOM EXP MEN [Paragraphs(11)]
    .insertparagraphafter
End With

With documento.Paragraphs(10).Range
    .Style = "Normal"
End With

' Insertar t�tulo de gr�fica NOM EXP MEN
With documento.Content
    .InsertAfter Hoja9.Cells(11, 2).Value ' T�tulo de gr�fica NOM EXP MEN [Paragraphs(12)]
    .insertparagraphafter
End With

With documento.Paragraphs(11).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica NOM EXP MEN
Grafica_nom.Chart.ChartArea.Copy
documento.Paragraphs(12).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(14)]
    .insertparagraphafter
End With

With documento.Paragraphs(13).Range
    .Style = "Fuentes"
End With

' Insertar salto de p�gina
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

' Insertar t�tulo de gr�fica TOT EXP ACU
With documento.Content
    .InsertAfter Hoja9.Cells(15, 2).Value ' T�tulo de gr�fica TOT EXP ACU [Paragraphs(18)]
    .insertparagraphafter
End With

With documento.Paragraphs(17).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica TOT EXP ACU
Grafica_acu.Chart.ChartArea.Copy
documento.Paragraphs(18).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(20)]
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
' Insertar salto de p�gina
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

' Insertar t�tulo de gr�fica TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(20, 2).Value ' T�tulo de gr�fica TOT EST [Paragraphs(25)]
    .insertparagraphafter
End With

With documento.Paragraphs(24).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica TOT EST
Grafica_tot_est.Chart.ChartArea.Copy
documento.Paragraphs(25).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(27)]
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

' Insertar salto de p�gina
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

' Insertar t�tulo de gr�fica DIST TOT EST
With documento.Content
    .InsertAfter Hoja9.Cells(25, 2).Value ' T�tulo de la gr�fica DIST TOT EST [Paragraphs(32)]
    .insertparagraphafter
End With

With documento.Paragraphs(31).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica DIST TOT EST
Grafica_dis_tot_est.Chart.ChartArea.Copy
documento.Paragraphs(32).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(34)]
    .insertparagraphafter
End With

With documento.Paragraphs(33).Range
    .Style = "Fuentes"
End With

' Insertar salto de p�gina
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


' Insertar t�tulo de gr�fica TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(30, 2).Value ' T�tulo de la gr�fica TOT TRA [Paragraphs(39)]
    .insertparagraphafter
End With

With documento.Paragraphs(38).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica TOT TRA
Grafica_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(39).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(41)]
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

' Insertar salto de p�gina
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

' Insertar t�tulo de gr�fica VAR TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(35, 2).Value ' T�tulo de la gr�fica VAR TOT TRA [Paragraphs(46)]
    .insertparagraphafter
End With

With documento.Paragraphs(45).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica VAR TOT TRA
Grafica_var_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(46).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(48)]
    .insertparagraphafter
End With

With documento.Paragraphs(47).Range
    .Style = "Fuentes"
End With

' Insertar salto de p�gina
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

' Insertar t�tulo de gr�fica DIST TOT TRA
With documento.Content
    .InsertAfter Hoja9.Cells(39, 2).Value ' T�tulo de la gr�fica DIST TOT TRA [Paragraphs(52)]
    .insertparagraphafter
End With

With documento.Paragraphs(51).Range
    .Style = "Figura - titulos"
End With

' Pasar gr�fica DIST TOT TRA
Grafica_dist_tot_tra.Chart.ChartArea.Copy
documento.Paragraphs(52).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con informaci�n de INEGI." ' Nota [Paragraphs(54)]
    .insertparagraphafter
End With

With documento.Paragraphs(53).Range
    .Style = "Fuentes"
End With


' Cambiar la fecha de realizaci�n
Set cuadrofecha = documento.Sections(1).Headers(1).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  350, 48 - 7, 240, 37.5 / 2)
                  ' wdHeaderFooterPrimary = 1
cuadrofecha.TextFrame.TextRange.Text = "Ficha informativa, " & diahoy & " de " & meshoypal & " de " & a�ohoy
cuadrofecha.TextFrame.TextRange.Font.Color = RGB(98, 113, 120)
cuadrofecha.TextFrame.TextRange.Font.Underline = wdUnderlineSingle
cuadrofecha.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
cuadrofecha.Fill.ForeColor = RGB(255, 255, 255)
cuadrofecha.Line.ForeColor = RGB(255, 255, 255)

' Guardar el documento (escribir la direcci�n en donde se quiera guardar)
documento.SaveAs "C:\Users\arturo.carrillo\Documents\IMMEX\" & a�obas & " " & mesbas & "\Ficha informativa IMMEX " & mesbaspal & " " & a�obas & "-" & a�ohoy & meshoy & diahoy & ".docx" ' UBICACI�N PERSONAL

End Sub

