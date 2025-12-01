Function CentrarTextEntreMarquesAlineacio()
    Dim doc As Document
    Dim rngDoc As Range
    Dim rngStart As Range
    Dim rngEnd As Range
    Dim rngInside As Range
    Dim found As Boolean
    Set doc = ActiveDocument
    Set rngDoc = doc.Content

    etiquetaInici = "#ALINEACIO#"
    etiquetaFinal = "#FIALINEACIO#"
    
    Application.ScreenUpdating = False
    Do
        ' 1) Buscar #ALINEACIO#
        Set rngStart = rngDoc.Duplicate
        With rngStart.Find
            .ClearFormatting
            .Text = etiquetaInici
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            found = .Execute
        End With
        If Not found Then Exit Do
        ' 2) Buscar #FIALINEACIO# després de #ALINEACIO#
        Set rngEnd = doc.Range(rngStart.End, doc.Content.End)
        With rngEnd.Find
            .ClearFormatting
            .Text = etiquetaFinal
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            found = .Execute
        End With
        If Not found Then Exit Do
        ' 3) Rang interior
        Set rngInside = doc.Range(rngStart.End, rngEnd.Start)
        ' Justificar paràgrafs interiors
        rngInside.ParagraphFormat.Alignment = wdAlignParagraphJustify
        ' 4) Esborrar marques
        rngEnd.Delete
        rngStart.Delete
        ' 5) Continuar amb la resta del document
        Set rngDoc = doc.Range(rngInside.End, doc.Content.End)
    Loop
    Application.ScreenUpdating = True
End Function
Function optimitzar_taules_tots_rangs() As Boolean
    'Optimitza les taules en TOTS els rangs marcats amb #INICI_AJUST_TAULES# i #FI_AJUST_TAULES#
    'Processa múltiples parelles d'etiquetes i després elimina totes les etiquetes del document
    
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim searchRange As Range
    Dim rngFinal As Range
    Dim rngZona As Range
    Dim taula As Table
    Dim etiquetaInici As String
    Dim etiquetaFinal As String
    Dim taulesAjustades As Long
    Dim rangsProcessats As Long
    Dim startZona As Long
    
    etiquetaInici = "#INICI_AJUST_TAULES#"
    etiquetaFinal = "#FI_AJUST_TAULES#"
    
    Set doc = ActiveDocument
    Set searchRange = doc.Content
    
    Application.ScreenUpdating = False
    
    Do
        'Cercar etiqueta d'inici a partir de la posició actual del searchRange
        With searchRange.Find
            .ClearFormatting
            .Text = etiquetaInici
            .MatchCase = False
            .Forward = True
            .Wrap = wdFindStop
            If Not .Execute Then Exit Do   'no hi ha més INICI
        End With
        
        'El rang trobat conté l'etiqueta d'inici
        startZona = searchRange.End      'després de l'etiqueta d'inici
        
        'Cercar etiqueta final a partir d'aquí
        Set rngFinal = doc.Range(startZona, doc.Content.End)
        With rngFinal.Find
            .ClearFormatting
            .Text = etiquetaFinal
            .MatchCase = False
            .Forward = True
            .Wrap = wdFindStop
            If Not .Execute Then Exit Do  'no hi ha FI ? sortim
        End With
        
        'Validem que el FI va després de l'INICI
        If rngFinal.Start <= startZona Then Exit Do
        
        'Rang de treball: entre les etiquetes (sense incloure-les)
        Set rngZona = doc.Range(startZona, rngFinal.Start)
        
        'Només taules dins del rang
        For Each taula In rngZona.Tables
            taula.AutoFitBehavior wdAutoFitWindow
            'Si realment necessites recalcular fórmules dins la taula, descomenta:
            'taula.Range.Calculate
            taulesAjustades = taulesAjustades + 1
        Next taula
        
        rangsProcessats = rangsProcessats + 1
        
        'Avancem el searchRange per continuar buscant més INICI més endavant
        searchRange.Start = rngFinal.End
        searchRange.End = doc.Content.End
        
    Loop
    
    'Eliminar totes les etiquetes d'inici
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = etiquetaInici
        .Replacement.Text = ""
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    'Eliminar totes les etiquetes de fi
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = etiquetaFinal
        .Replacement.Text = ""
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    
    optimitzar_taules_tots_rangs = (rangsProcessats > 0)
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Function
    
ErrorHandler:
    optimitzar_taules_tots_rangs = False
    Resume CleanExit
End Function
Function SubstituirSpaces()
    Dim numSubstitucions As Long
    'Dim totalSubstitucions As Long
    Dim taula As Table
    Dim i As Integer
	Dim etiquetaSpace As String
	
	etiquetaSpace = "ø"
    
    ' Desactiva l'actualització de pantalla per millorar el rendiment
    Application.ScreenUpdating = False
    
    'totalSubstitucions = 0
    
    ' 1. SUBSTITUCIÓ EN EL COS PRINCIPAL DEL DOCUMENT
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' Configura la cerca
        .Text = etiquetaSpace
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        
        ' Executa la substitució de totes les ocurrències
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Compta les substitucions del cos principal
    'totalSubstitucions = ContarOcurrencies(ActiveDocument.Content)
    
    ' 2. SUBSTITUCIÓ EN TOTES LES TAULES
    For Each taula In ActiveDocument.Tables
        With taula.Range.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            
            ' Configura la cerca per a la taula
            .Text = etiquetaSpace
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            
            ' Executa la substitució en aquesta taula
            .Execute Replace:=wdReplaceAll
        End With
        
        ' Suma les substitucions d'aquesta taula
        'totalSubstitucions = totalSubstitucions + ContarOcurrencies(taula.Range)
    Next taula
    
    ' Reactiva l'actualització de pantalla
    Application.ScreenUpdating = True
    
    ' Mostra un missatge amb el resultat
    'If ActiveDocument.Tables.Count > 0 Then
        'MsgBox "S'han substituït un total de " & totalSubstitucions & " ocurrències de #SPACE# per espais en blanc" & vbCrLf & _
        '       "(incloent " & ActiveDocument.Tables.Count & " taula/es).", vbInformation, "Substitució completada"
    'Else
        'MsgBox "S'han substituït " & totalSubstitucions & " ocurrències de #SPACE# per espais en blanc.", vbInformation, "Substitució completada"
    'End If
    
End Function

Function formatar_text_entre_etiquetes(nom_font As String, mida_font As Integer) As Boolean
    'Formata el text contingut entre {FORMAT_INI} i {FORMAT_FIN}
    'Versió millorada que gestiona salts de línia i espais
    
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim posInici As Long
    Dim posFinal As Long
    Dim comptador As Integer
    Dim etiquetaInici As String
    Dim etiquetaFinal As String
    
    etiquetaInici = "{FORMAT_INI}"
    etiquetaFinal = "{FORMAT_FIN}"
    comptador = 0
    
    'Cerquem l'etiqueta d'inici
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Text = etiquetaInici
        .MatchCase = False
        .MatchWholeWord = False
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            'Guardem la posició inicial (després de l'etiqueta)
            posInici = rng.End
            
            'Ara cerquem l'etiqueta final des d'aquesta posició
            Set rngFinal = ActiveDocument.Range(posInici, ActiveDocument.Content.End)
            
            With rngFinal.Find
                .ClearFormatting
                .Text = etiquetaFinal
                .MatchCase = False
                .MatchWholeWord = False
                .Forward = True
                .Wrap = wdFindStop
                
                If .Execute Then
                    'Hem trobat la parella d'etiquetes
                    posFinal = rngFinal.Start
                    
                    'Creem un rang que inclou des de l'inici de FORMAT_INI fins al final de FORMAT_FIN
                    Set rngComplet = ActiveDocument.Range(rng.Start, rngFinal.End)
                    
                    'Creem un rang només amb el text (sense etiquetes)
                    Set rngText = ActiveDocument.Range(posInici, posFinal)
                    
                    'Guardem el text net
                    Dim textNet As String
                    textNet = rngText.Text
                    
                    'Reemplacem tot el bloc (etiquetes + text) amb només el text
                    rngComplet.Text = textNet
                    
                    'Apliquem el format al text net
                    With rngComplet.Font
                        .Name = nom_font
                        .Size = mida_font
                    End With
                    
                    comptador = comptador + 1
                End If
            End With
            
            'Continuem buscant des de després del rang actual
            Set rng = ActiveDocument.Range(rngComplet.End, ActiveDocument.Content.End)
        Loop
    End With
    
    If comptador > 0 Then
        'MsgBox "S'han formatat " & comptador & " blocs de text.", vbInformation
        formatar_text_entre_etiquetes_v3 = True
    Else
        'MsgBox "No s'ha trobat cap text entre " & etiquetaInici & " i " & etiquetaFinal & ".", vbInformation
        formatar_text_entre_etiquetes_v3 = False
    End If
    
    Exit Function
    
ErrorHandler:
    'MsgBox "Error formatant el text: " & Err.Description, vbCritical
    formatar_text_entre_etiquetes_v3 = False
End Function
Function optimitzar_totes_columnes_finestra() As Boolean
    'Optimitza totes les taules del document en funció de la finestra
    'Retorna True si s'han pogut ajustar totes, False si hi ha hagut algun error
    
    On Error GoTo ErrorHandler
    
    Dim taula As Table
    Dim totalTaules As Integer
    Dim taulesAjustades As Integer
    
    totalTaules = ActiveDocument.Tables.Count
    
    'Validem que hi hagi taules al document
    If totalTaules = 0 Then
        'MsgBox "El document no conté cap taula.", vbInformation
        optimitzar_totes_columnes_finestra = False
        Exit Function
    End If
    
    taulesAjustades = 0
    
    'Iterem per totes les taules del document
    For Each taula In ActiveDocument.Tables
        taula.AutoFitBehavior wdAutoFitWindow
        taula.Range.Calculate 'Actualitzem la vista
        taulesAjustades = taulesAjustades + 1
    Next taula
    
    'Missatge informatiu (opcional, pots comentar-lo si no el vols)
    'MsgBox "S'han ajustat " & taulesAjustades & " taules a la finestra.", vbInformation
    
    optimitzar_totes_columnes_finestra = True
    Exit Function
    
ErrorHandler:
    'MsgBox "Error ajustant les taules. S'han ajustat " & taulesAjustades & _
           " de " & totalTaules & " taules." & vbCrLf & _
           "Error: " & Err.Description, vbCritical
    optimitzar_totes_columnes_finestra = False
End Function
Function PosarNegretaEntreCaracters()
    On Error Resume Next
    
    Dim doc As Document
    Dim rng As Range
    Dim posInici As Long
    Dim posFi As Long
    Dim textCerca As String
    
    ' Intentar obtenir el document actiu
    Set doc = ActiveDocument
    If doc Is Nothing Then
        ' MsgBox "No s'ha pogut obtenir el document", vbCritical
        Exit Function
    End If
    
    ' MsgBox "Inici macro - Document: " & doc.Name, vbInformation
    
    ' Desactivar actualització de pantalla per millorar rendiment
    Application.ScreenUpdating = False
    
    ' Buscar i processar cada ocurrència
    Do
        Set rng = doc.Content
        rng.Find.ClearFormatting
        
        With rng.Find
            .Text = "{B}"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            
            If .Execute Then
                posInici = rng.Start
                ' MsgBox "Trobat {B} a posició: " & posInici, vbInformation
                
                ' Buscar el tancament
                Set rng = doc.Range(Start:=posInici + 3, End:=doc.Content.End)
                rng.Find.ClearFormatting
                rng.Find.Text = "{/B}"
                rng.Find.Forward = True
                rng.Find.Wrap = wdFindStop
                
                If rng.Find.Execute Then
                    posFi = rng.Start
                    ' MsgBox "Trobat {/B} a posició: " & posFi, vbInformation
                    
                    ' Seleccionar el text entre etiquetes
                    Set rng = doc.Range(Start:=posInici + 3, End:=posFi)
                    rng.Font.Bold = True
                    
                    ' Eliminar {/B}
                    Set rng = doc.Range(Start:=posFi, End:=posFi + 4)
                    rng.Delete
                    
                    ' Eliminar {B}
                    Set rng = doc.Range(Start:=posInici, End:=posInici + 3)
                    rng.Delete
                    
                    ' MsgBox "Etiquetes processades correctament", vbInformation
                Else
                    ' MsgBox "No s'ha trobat {/B} corresponent", vbExclamation
                    Exit Do
                End If
            Else
                ' MsgBox "No hi ha més etiquetes {B}", vbInformation
                Exit Do
            End If
        End With
    Loop
    
    Application.ScreenUpdating = True
    ' MsgBox "Macro finalitzada correctament", vbInformation
End Function

Function Obtenir_Llista_de_Vincles()
   Dim R3Table As Object
   Dim Row     As Object
   Dim aField  As Field
   Dim oSection As Section
   Dim oHeader As HeaderFooter
   Dim oRng As Range
   
   Set R3Table = ActiveDocument.Container.Tables("LINK_LIST").Table

   R3Table.Rows.RemoveAll
   'Leemos los SCRIPTS del cuerpo del documento
      For Each aField In ActiveDocument.Fields
        Set Row = R3Table.Rows.Add
        Row.Cell(1) = aField.Code.Text
   Next aField

'Leemos los SCRIPTS de la cabecera
      For Each oSection In ActiveDocument.Sections
        For Each oHeader In oSection.Headers
        Set oRng = oHeader.Range
            For Each aField In oRng.Fields
            Set Row = R3Table.Rows.Add
            Row.Cell(1) = aField.Code.Text
            Next aField
        Next oHeader
      Next oSection
      
'Leemos los SCRIPTS del pie de página
      For Each oSection In ActiveDocument.Sections
        For Each oHeader In oSection.Footers
        Set oRng = oHeader.Range
            For Each aField In oRng.Fields
            Set Row = R3Table.Rows.Add
            Row.Cell(1) = aField.Code.Text
            Next aField
        Next oHeader
      Next oSection

End Function
Function Obtenir_Llista_de_Taules()
   Dim R3Table  As Object
   Dim aTable   As Table
   Dim aCounter As Integer
 
   Set R3Table = ActiveDocument.Container.Tables("TABLE_LIST").Table
   R3Table.Rows.RemoveAll
 
   aCounter = 1
   For Each aTable In ActiveDocument.Tables
      Call analitza_taula(aCounter, aTable, R3Table)
      aCounter = aCounter + 1
   Next aTable
End Function
Function analitza_taula(aCounter As Integer, aTable As Table, R3Table As Object)
  Dim aRow  As Row
  Dim aCell As Cell
  Dim Row   As Object
  Dim colCount As Integer
  Dim rowCount As Integer
    
  If aTable.Uniform = False Then
    Exit Function
  End If
  
  rowCount = 1
  For Each aRow In aTable.Rows
    colCount = 1
    For Each aCell In aRow.Cells
      If aCell.Range.Text <> "" Then
        Set Row = R3Table.Rows.Add
        Row.Cell(1) = aCounter
        Row.Cell(2) = aCell.Range.Text
        Row.Cell(3) = colCount
        Row.Cell(4) = rowCount
      End If
      colCount = colCount + 1
    Next aCell
    rowCount = rowCount + 1
  Next aRow
  
End Function
Function EliminarTablaSiSegundaFilaVacia()
    Dim tbl As Table
    Dim celda As Cell
    Dim filaVacia As Boolean
    Dim contenidoCelda As String
    
    For Each tbl In ActiveDocument.Tables
        If tbl.Rows.Count = 2 Then
            filaVacia = True
            For Each celda In tbl.Rows(2).Cells
                contenidoCelda = Trim(celda.Range.Text)
                contenidoCelda = Replace(contenidoCelda, Chr(13) & Chr(7), "")
                If contenidoCelda <> "" Then
                    filaVacia = False
                    Exit For
                End If
            Next celda
            If filaVacia Then
                tbl.Delete
            End If
        End If
    Next tbl
End Function

Function elimina_columnes_no_usades(id_taula As Integer)
    'Elimina totes les columnes que estiguin en blanc
    Dim columna As Column
    Dim substring As String
    'Recorrem totes les columnes de la taula indicada per paràmetre
    For Each columna In ActiveDocument.Tables(id_taula).Columns
        ' Seleccionem el text de la primera columna
        With columna.Cells(1).Range
            .MoveEnd Unit:=wdCharacter, Count:=-1
            ' Si la cel·la està buida o té el caràcter per defecte "[" eliminem tota la columna
            substring = Left(.Text, 2)
            If substring = "[i" Or .Text = "" Then
                columna.Delete
            End If
        End With
        Next columna
End Function
Function optimitzar_columnes_contingut(id_taula As Integer)
    'Optimitzem la taula indicada per paràmetre en funció del seu contingut
    'ActiveDocument.Tables(id_taula).AutoFitBehavior _
    '    wdAutoFitContent
    Dim aCounter As Integer
    aCounter = 1
    For Each aTable In ActiveDocument.Tables
        If aCounter = id_taula Then
            ActiveDocument.Tables(id_taula).AutoFitBehavior _
                wdAutoFitContent
        End If
        aCounter = aCounter + 1
    Next aTable
End Function
Function optimitzar_columnes_finestra(id_taula As Integer) As Boolean
    'Optimitzem la taula indicada per paràmetre en funció de la finestra
    'Retorna True si s'ha pogut ajustar, False si hi ha hagut algun error
    
    On Error GoTo ErrorHandler
    
    'Validem que l'ID de taula sigui vàlid
    If id_taula < 1 Or id_taula > ActiveDocument.Tables.Count Then
        'MsgBox "Error: La taula " & id_taula & " no existeix. El document té " & _
               ActiveDocument.Tables.Count & " taules.", vbExclamation
        optimitzar_columnes_finestra = False
        Exit Function
    End If
    
    'Ajustem directament la taula indicada
    ActiveDocument.Tables(id_taula).AutoFitBehavior wdAutoFitWindow
    
    'Opcionalment, també podem assegurar que s'actualitzi la vista
    ActiveDocument.Tables(id_taula).Range.Calculate
    
    optimitzar_columnes_finestra = True
    Exit Function
    
ErrorHandler:
    'MsgBox "Error ajustant la taula " & id_taula & ": " & Err.Description, vbCritical
    optimitzar_columnes_finestra = False
End Function
Public Function get_si_plantilla_plec_dinamic() As Integer
get_si_plantilla_dinamica = 1
End Function
Function optimitzar_columnes_taula()
    Dim aTable   As Table
    For Each aTable In ActiveDocument.Tables
        aTable.AutoFitBehavior _
            wdAutoFitContent
    Next aTable
End Function
Function Actualitza_Vincles()
   Dim story As Word.Range

'Actualización de Scripts del cuerpo del documento
ActiveDocument.Fields.Update
'Actualización de tablas
ActiveDocument.Fields.Unlink

'Actualización de SCRIPTS en cabeceras y pie de páginas
    For Each story In ActiveDocument.StoryRanges
        Do
            story.Fields.Update
            Set story = story.NextStoryRange
        Loop Until (story Is Nothing)
    Next
End Function
Function eliminarColumnesAmbBuides()
    Dim aTable As Table
    Dim i As Long, j As Long
    Dim colToDelete As Collection
    Dim cel As Cell
    Dim textCel As String
    Dim idx As Long
    
    On Error GoTo ManejarError
    
    For Each aTable In ActiveDocument.Tables
        Set colToDelete = New Collection
        
        ' Recollir columnes a eliminar segons cel·les amb "[i"
        For Each cel In aTable.Range.Cells
            textCel = cel.Range.Text
            textCel = Left(textCel, Len(textCel) - 2) ' Treure caràcters finals de control
            If InStr(textCel, "[i") > 0 Then
                On Error Resume Next
                colToDelete.Add cel.ColumnIndex, CStr(cel.ColumnIndex) ' Evitar duplicats
                On Error GoTo 0
            End If
        Next cel
        
        ' Eliminar columnes de darrere a davant per evitar canvis d'índex
        For idx = colToDelete.Count To 1 Step -1
            Dim colIndex As Long
            colIndex = colToDelete(idx)
            ' Eliminar cel·les de la columna (fila per fila)
            For i = aTable.Rows.Count To 1 Step -1
                On Error Resume Next
                aTable.Cell(i, colIndex).Delete
                On Error GoTo 0
            Next i
        Next idx
    Next aTable
    
    Exit Function
    
ManejarError:
    MsgBox "Error: " & Err.Description
    Resume Next
End Function