Attribute VB_Name = "Módulo1"
Option Explicit

Sub CREA_CT()

    ' --- DECLARACIÓN DE VARIABLES ---
    Dim wordApp As Object
    Dim doc As Object
    Dim hoja As Worksheet
    Dim CT As Worksheet
    Dim fila As Long ' Aunque no se usa directamente en este código, se mantiene.
    Dim rutaPlantilla As String, rutaSalida As String
    
    ' Variables para los condicionales
    Dim ManometroAnalogico As String
    Dim ManometroElectronico As String
    Dim longitudMM As String
    Dim diamCol As Variant
    Dim diamEle As Variant
    Dim contador As String ' Aunque comentado, se mantiene la declaración.
    Dim AnguloTlv As String
    Dim materialJaula As String
    Dim espTolvaValue As Variant
    Dim mater_CAL As String
    Dim mater_CAS As String
    Dim mater_CHV As String
    Dim mater_TLV As String
    
    On Error GoTo ErrorHandler

    ' --- CONFIGURACIÓN DE HOJAS Y RUTAS ---
    Set hoja = ThisWorkbook.Sheets("CT")
    fila = 2 ' No utilizada en el código actual, pero se mantiene la inicialización.
    
    ' Define la ruta de la plantilla. Descomenta la que necesites.
    rutaPlantilla = "Y:\COSTES\PLANTILLAS\CARACTERISTICAS TECNICAS.docx"
    
    'rutaPlantilla = "C:\Users\Christian Rossi\PRUEBAS\CT_excel\CARACTERISTICAS TECNICAS.docx" ' Ruta de ejemplo para pruebas
    
    ' Crea el nombre del archivo de salida de forma dinámica
    rutaSalida = ThisWorkbook.Path & "\CT-" & hoja.Range("B1").Value & ".docx"
    'rutaSalida = "C:\Users\Christian Rossi\PRUEBAS\CT_excel\CT-" & hoja.Range("E6").Value & ".docx" ' Ruta de ejemplo para pruebas
    
    ' Verifica si la plantilla existe
    If Dir(rutaPlantilla) = "" Then
        MsgBox "No se encontró la plantilla: " & rutaPlantilla, vbCritical
        Exit Sub
    End If

    ' --- INICIAR WORD ---
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add(rutaPlantilla)
    
    ' --- LÓGICA CONDICIONAL: ASIGNACIÓN DE VALORES ---
    ' Condicional para Manómetros - Usando Select Case para mayor claridad
    Select Case hoja.Range("B30").Value
        Case 0, 1, 2, 3, 4, 5, 6
            ManometroAnalogico = "Esfera 0-300 mmca"
            ManometroElectronico = "Timer-manómetro. 220 Vac / IP56 / 50 Hz "
        Case Else
            ManometroAnalogico = "-"
            ManometroElectronico = "Timer-manómetro. 220 Vac / IP56 / 50 Hz / signal 4-20 mA"
    End Select
    
    ' Ángulo tolva 60 / 70 º
    If hoja.Range("B22").Value = "2" Then
        AnguloTlv = "70"
    Else
        AnguloTlv = "60"
    End If
    
    ' Condicional para longitud de las mangas
    Select Case hoja.Range("B25").Value
        Case 4: longitudMM = "1.261"
        Case 6: longitudMM = "1.870"
        Case 8: longitudMM = "2.479"
        Case 10: longitudMM = "3.088"
        Case 12: longitudMM = "3.697"
        Case Else: longitudMM = "0"
    End Select
    
    ' Condicional para los diámetros
    Select Case hoja.Range("B27").Value
        Case 2, 3, 8, 9
            diamCol = 6
            diamEle = 1
        Case 4, 5, 10, 11
            diamCol = 8
            diamEle = "1 1/2"
        Case 6, 7, 12, 13
            diamCol = 8
            diamEle = 2
        Case Else
            diamCol = ""
            diamEle = ""
    End Select
    
    ' Obtiene los primeros 2 caracteres del nombre del archivo para el contador
    ' Esta linea fue comentada por ti, la mantengo asi
    ' contador = Left(ThisWorkbook.Name, 2)
    
    ' --- LÓGICA CONDICIONAL PARA EL MATERIAL DE LA JAULA ---
    materialJaula = hoja.Range("B26").Value
    If materialJaula = "Pintadas" Then
        materialJaula = "Acero pintado"
    End If
    ' --- FIN LÓGICA CONDICIONAL PARA EL MATERIAL DE LA JAULA ---

    ' --- LÓGICA CONDICIONAL PARA MATER_CAL ---
    mater_CAL = "" ' Inicializa la variable por seguridad
    Select Case hoja.Range("B18").Value
        Case 1: mater_CAL = "S235JR"
        Case 2: mater_CAL = "AISI-304"
        Case 3: mater_CAL = "AISI-316"
        Case Else: mater_CAL = "" ' Valor por defecto si no coincide
    End Select
    ' --- FIN LÓGICA CONDICIONAL PARA MATER_CAL ---
    
    ' --- LÓGICA CONDICIONAL PARA MATER_CAS ---
    mater_CAS = "" ' Inicializa la variable por seguridad
    Select Case hoja.Range("B20").Value
        Case 1: mater_CAS = "S235JR"
        Case 2: mater_CAS = "AISI-304"
        Case 3: mater_CAS = "AISI-316"
        Case Else: mater_CAS = "" ' Valor por defecto si no coincide
    End Select
    ' --- FIN LÓGICA CONDICIONAL PARA MATER_CAS ---
    
 ' --- LÓGICA CONDICIONAL PARA ESP_TOLVA Y MATER_TLV ---
    ' Inicializa mater_TLV por seguridad antes de la lógica condicional
    mater_TLV = ""
    
    ' Unificamos la lógica para espTolvaValue y mater_TLV
    If (hoja.Range("B31").Value) = "A" Or (hoja.Range("B31").Value) = "AE" Or (hoja.Range("B31").Value) = "PL" Then
        espTolvaValue = "-"
        mater_TLV = "-" ' Si C62 es A, AE o PL, ambos serán "-"
    Else
        espTolvaValue = hoja.Range("B17").Value
        ' Si NO es "A", "AE" o "PL", entonces mater_TLV se define por CT.Range("M137").Value
        Select Case hoja.Range("B21").Value
            Case 1: mater_TLV = "S235JR"
            Case 2: mater_TLV = "AISI-304"
            Case 3: mater_TLV = "AISI-316"
            Case Else: mater_TLV = "" ' Valor por defecto si no coincide
        End Select
    End If
    ' --- FIN LÓGICA CONDICIONAL PARA ESP_TOLVA Y MATER_TLV ---
    
    ' --- LÓGICA CONDICIONAL PARA MATER_CHV ---
    mater_CHV = "" ' Inicializa la variable por seguridad
    Select Case hoja.Range("B19").Value
        Case 1: mater_CHV = "S235JR"
        Case 2: mater_CHV = "AISI-304"
        Case 3: mater_CHV = "AISI-316"
        Case Else: mater_CHV = "" ' Valor por defecto si no coincide
    End Select
    ' --- FIN LÓGICA CONDICIONAL PARA MATER_CHV ---

    ' --- MODIFICACIONES PARA CAMPOS ESPECÍFICOS ---
    ' Modificación para BAR: Si CT.Range("BY3") es un error (#N/A, #DIV/0!, etc.), usa "XXXXXXX"
    If IsError(hoja.Range("B32").Value) Then
        Call ReemplazarCampo(doc, "{{BAR}}", "XXXXXXX")
    Else
        Call ReemplazarCampo(doc, "{{BAR}}", hoja.Range("B32").Value)
    End If

    ' Modificación para CONSUMO_AIRE: Si es numérico, lo formatea; si no, usa "XX"
    If IsNumeric(hoja.Range("B11").Value) Then
        Call ReemplazarCampo(doc, "{{CONSUMO_AIRE}}", Format(hoja.Range("B11").Value, "0"))
    Else
        Call ReemplazarCampo(doc, "{{CONSUMO_AIRE}}", "XX")
    End If

    ' --- REEMPLAZAR TODOS LOS CAMPOS EN EL CUERPO PRINCIPAL ---
    Call ReemplazarCampo(doc, "{{NOF}}", hoja.Range("B1").Value)
    Call ReemplazarCampo(doc, "{{CAUDAL}}", Format(hoja.Range("B3").Value, "#,##0"))
    Call ReemplazarCampo(doc, "{{PRODUCT}}", hoja.Range("B5").Value)
    Call ReemplazarCampo(doc, "{{CONC}}", IIf(IsEmpty(hoja.Range("B6").Value) Or hoja.Range("B6").Value = 0, "20÷30", hoja.Range("B6").Value))
    Call ReemplazarCampo(doc, "{{DENS}}", IIf(IsEmpty(hoja.Range("B7").Value) Or hoja.Range("B7").Value = 0, "800", hoja.Range("B7").Value))
    'Call ReemplazarCampo(doc, "{{TEMPERATURA}}", hoja.Range("B10").Value)
    Call ReemplazarCampo(doc, "{{TEMPERATURA}}", IIf(IsEmpty(hoja.Range("B4").Value), "20", hoja.Range("B4").Value))
    Call ReemplazarCampo(doc, "{{FILTRO}}", hoja.Range("B8").Value)
    Call ReemplazarCampo(doc, "{{SUP_FILTRANTE}}", hoja.Range("B9").Value)
    Call ReemplazarCampo(doc, "{{RATIO_F}}", Format(hoja.Range("B10").Value, "0.00"))
    Call ReemplazarCampo(doc, "{{P_T}}", hoja.Range("B12").Value)
    Call ReemplazarCampo(doc, "{{P_d}}", hoja.Range("B13").Value)
    Call ReemplazarCampo(doc, "{{ESP_CAL}}", hoja.Range("B14").Value)
    Call ReemplazarCampo(doc, "{{ESP_VENT}}", hoja.Range("B15").Value)
    Call ReemplazarCampo(doc, "{{ESP_CAS}}", hoja.Range("B16").Value)
    Call ReemplazarCampo(doc, "{{ESP_TOLVA}}", espTolvaValue)
    Call ReemplazarCampo(doc, "{{ANG}}", AnguloTlv)
    Call ReemplazarCampo(doc, "{{MAT_MANGA}}", hoja.Range("B23").Value)
    Call ReemplazarCampo(doc, "{{NUM_MANGAS}}", hoja.Range("B24").Value)
    Call ReemplazarCampo(doc, "{{LONG_MANGA}}", longitudMM)
    Call ReemplazarCampo(doc, "{{MAT_JAULA}}", materialJaula)
    Call ReemplazarCampo(doc, "{{NUM_VALV}}", hoja.Range("B28").Value)
    Call ReemplazarCampo(doc, "{{TIMER_MAN_CT}}", ManometroElectronico)
    Call ReemplazarCampo(doc, "{{MAN_CT}}", ManometroAnalogico)
    Call ReemplazarCampo(doc, "{{DIAM_COL}}", diamCol)
    Call ReemplazarCampo(doc, "{{DIAM_ELE}}", diamEle)
    Call ReemplazarCampo(doc, "{{MATER_CAL}}", mater_CAL)
    Call ReemplazarCampo(doc, "{{MATER_CAS}}", mater_CAS)
    Call ReemplazarCampo(doc, "{{MATER_TLV}}", mater_TLV) ' Uso único y correcto de mater_TLV
    Call ReemplazarCampo(doc, "{{MATER_CHV}}", mater_CHV)
    ' Esta linea fue comentada por ti, la mantengo asi
    ' Call ReemplazarCampo(doc, "{{CONTADOR}}", contador)
    
    ' --- AÑADIR NOF AL FOOTER ---
    Dim seccion As Object
    Dim pieDePagina As Object
    Dim i As Long
    
    For Each seccion In doc.Sections
        For i = 1 To seccion.Footers.Count
            Set pieDePagina = seccion.Footers(i)
            With pieDePagina.Range.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "{{NOF}}"
                .Replacement.Text = CStr(hoja.Range("B1").Value) ' Convertir a String para evitar problemas
                .Replacement.Font.Name = "Arial"
                .Replacement.Font.Size = 9
                .Execute Replace:=2
            End With
            
            ' --- ACTUALIZAR CAMPOS DEL FOOTER ---
            pieDePagina.Range.Fields.Update
            ' -----------------------------------
        Next i
    Next seccion

    ' --- GUARDAR DOCUMENTO Y FINALIZAR ---
    doc.SaveAs2 rutaSalida, FileFormat:=16 ' FileFormat:=16 para .docx
    MsgBox "Documento completado guardado en:" & vbCrLf & rutaSalida, vbInformation

    Exit Sub ' Salir de la subrutina si todo fue exitoso

ErrorHandler:
    ' Mensaje de error más descriptivo
    MsgBox "Se produjo un error al crear el documento: " & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical
    
    ' Cerrar el documento y la aplicación Word si están abiertos
    If Not doc Is Nothing Then doc.Close False ' Cerrar sin guardar cambios
    If Not wordApp Is Nothing Then wordApp.Quit
    
    ' Limpiar objetos para liberar memoria
    Set doc = Nothing
    Set wordApp = Nothing
End Sub

' Esta subrutina aplica el formato Arial 9 a los campos del cuerpo principal
' y ahora convierte el valor a String antes de reemplazarlo.
Sub ReemplazarCampo(doc As Object, marcador As String, valor As Variant)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = marcador
        ' Usa CStr para convertir el valor a String, manejando correctamente errores y otros tipos
        .Replacement.Text = IIf(IsEmpty(valor), "", CStr(valor))
        
        .Replacement.Font.Name = "Arial"
        .Replacement.Font.Size = 9
        
        .Execute Replace:=2 ' Reemplazar todas las ocurrencias
    End With
End Sub


