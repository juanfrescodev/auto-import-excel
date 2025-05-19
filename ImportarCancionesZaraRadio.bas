Attribute VB_Name = "M�dulo1"
Sub ImportarCancionesZaraRadio(Optional modo As String = "")

    If LCase(modo) <> "auto" Then
        Exit Sub
    End If

    Dim logFile As String
    Dim line As String
    Dim row As Long
    Dim logFileNumber As Integer
    Dim currentDate As String
    Dim parts() As String
    Dim filePathParts() As String
    Dim fileName As String
    Dim songName As String
    Dim artistName As String
    Dim palabrasExcluidas As Variant
    Dim i As Integer
    Dim excluir As Boolean
    Dim rutaCompleta As String
    Dim horario As String
    Dim cancionesRegistradas As Object
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim existingSongs As Object
    Dim existingSongKey As String
    
    ' Lista de palabras clave para excluir archivos no deseados
    palabrasExcluidas = Array("Separador", "RNB", "minutos", "Art�stica", "Publicidad", "Recorte", "ID", "Promo", "Cortina", "Spot")
    
    ' Obtener la fecha actual en el formato A�o-Mes-D�a
    currentDate = Format(Date, "yyyy-mm-dd")
    
    ' Ruta del archivo de log de ZaraRadio (ajustar seg�n donde se guardan los logs)
    logFile = "C:\Logs\ZaraRadio\" & currentDate & ".log" ' AJUSTA ESTA RUTA
    
    ' Verificar si el archivo existe
    If Dir(logFile) = "" Then
        MsgBox "El archivo de log no se encuentra: " & logFile, vbExclamation
        Exit Sub
    End If
    
    ' Crear un diccionario para evitar duplicados
    Set cancionesRegistradas = CreateObject("Scripting.Dictionary")
    Set existingSongs = CreateObject("Scripting.Dictionary")
    
    ' Crear un objeto RegExp para manejar el nombre de la canci�n
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = "^([^-\(]+)\s*-\s*(.+)$"  ' Expresi�n regular para extraer Artista y Canci�n
    
    ' Abrir el archivo de log
    logFileNumber = FreeFile
    Open logFile For Input As logFileNumber

    ' Encontrar la primera fila vac�a en la hoja activa
    row = Cells(Rows.Count, 1).End(xlUp).row + 1

    ' Leer el archivo l�nea por l�nea
    Do While Not EOF(logFileNumber)
        Line Input #logFileNumber, line
        
        ' Verificar si la l�nea contiene "inicio" (indica una canci�n reproducida)
        If InStr(line, "inicio") > 0 Then
            ' Dividir la l�nea por TABULADOR
            parts = Split(line, vbTab)
            
            ' Verificar que la l�nea tiene al menos 3 partes
            If UBound(parts) >= 2 Then
                ' Obtener la ruta completa del archivo
                rutaCompleta = parts(2)
                
            If InStr(1, rutaCompleta, "\03 ARTISTICA\", vbTextCompare) > 0 Then
                ' Permitir solo si es la subcarpeta Servicio de mensajes
                If InStr(1, rutaCompleta, "\03 ARTISTICA\02. SERVICIO DE MENSAJES\", vbTextCompare) = 0 Then
                    GoTo Continuar
                End If
            End If
               
                ' Obtener el horario
                horario = Trim(parts(0))
                
                ' Dividir por "\" para obtener el nombre del archivo
                filePathParts = Split(rutaCompleta, "\")
                fileName = filePathParts(UBound(filePathParts)) ' �ltimo elemento es el archivo
                
                ' Quitar la extensi�n .mp3
                If InStr(fileName, ".mp3") > 0 Then
                    fileName = Replace(fileName, ".mp3", "")
                End If
                
                ' Filtrar archivos no deseados
                excluir = False
                For i = LBound(palabrasExcluidas) To UBound(palabrasExcluidas)
                    If InStr(1, fileName, palabrasExcluidas(i), vbTextCompare) > 0 Then
                        excluir = True
                        Exit For
                    End If
                Next i
                
                ' Si el archivo debe excluirse, saltamos a la siguiente iteraci�n
                If excluir Then GoTo Continuar
                
                ' Usar RegExp para dividir el nombre del archivo en artista y canci�n
                Set matches = regEx.Execute(fileName)
                
                ' Verificar si se encontr� una coincidencia
                If matches.Count > 0 Then
                    Set match = matches(0)
                    artistName = Trim(match.SubMatches(0))  ' Artista
                    songName = Trim(match.SubMatches(1))    ' Canci�n
                    
                    ' Generar una clave �nica para cada canci�n (Artista - Canci�n)
                    existingSongKey = songName & " - " & artistName
                    
                    ' Verificar si la canci�n ya ha sido registrada
                    If Not existingSongs.exists(existingSongKey) Then
                        ' Guardar en la hoja de Excel
                        Cells(row, 1).Value = currentDate
                        Cells(row, 2).Value = horario
                        Cells(row, 3).Value = songName
                        Cells(row, 4).Value = artistName
                        
                        ' Agregar la canci�n al diccionario para evitar duplicados
                        existingSongs.Add existingSongKey, True
                        
                        row = row + 1
                    End If
                End If
            End If
        End If
Continuar:
    Loop

    ' Cerrar el archivo de log
    Close logFileNumber
    
    ' Liberar memoria del diccionario
    Set cancionesRegistradas = Nothing
    Set existingSongs = Nothing
    Set regEx = Nothing
    
        ' Guardar autom�ticamente el archivo
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:="C:\Aire AM\09 Planilla AADI CAPIF\2025\05 mayo\adi.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

End Sub


