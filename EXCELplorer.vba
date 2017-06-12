Option Explicit

Sub EXCELplorer()

Dim txtFileName() As String
Dim s As Series
Dim x_temp(10000) As Double
Dim y_temp(10000) As Double
Dim MyChart As Chart
Dim lineArray() As String
Dim x_label() As String
Dim y_label() As String
Dim WordFileName As String
Dim fileNumber As Integer
Dim iSheet As Integer
Dim delim As String

delim = Application.DecimalSeparator
Application.DecimalSeparator = "."

' Selection du fichier Word dans lequel on souhaite sauvegarder les figures
Call selectWordFile(WordFileName)

' Si aucun fichier sélectionner la macro s'arrête
If (WordFileName = "") Then
    Exit Sub
End If

' On isole le nom du dossier dans lequel le fichier Word est. Ce dossier sera utilisé pour créer les fichiers Powerpoint temporaires
Dim directory As String
directory = Left(WordFileName, InStrRev(WordFileName, "\"))

' Lecture de tous les fichiers pour lesquels on veut générer les graphiques
Call openDialog(txtFileName, fileNumber)

' Si aucun fichier sélectionner la macro s'arrête
If fileNumber = 0 Then
    Exit Sub
End If

' On empêche Excel de faire de rafraîchir l'écran
Application.ScreenUpdating = False

' Tous les tracés sont faits dans la feuille "Graphics". Cette feuille est d'abord supprimée si elle existe puis recréée
Application.DisplayAlerts = False
For iSheet = 1 To Sheets.Count
    If Sheets(iSheet).Name = "Graphics" Then
        Sheets(iSheet).Delete
    End If
Next
Application.DisplayAlerts = True

Worksheets.Add
ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.Count)
ActiveSheet.Name = "Graphics"

ReDim x_label(fileNumber)
ReDim y_label(fileNumber)

Dim i_file As Integer
i_file = 0

Dim file As Variant
For Each file In txtFileName

    i_file = i_file + 1

    ' Création de l'objet qui va contenir les courbes
    ActiveSheet.Shapes.AddChart.Select
    Dim sh As Variant 'mais quel est le type de cette variable ?!
    Set sh = Selection

    ' Suppression de la bordure du cadre
    ActiveChart.ChartArea.Border.LineStyle = xlNone

    Dim nSeries As Integer
    nSeries = 0
    Open file For Input As #1

    Dim nLine As Integer
    For nLine = 1 To 7

        Dim textline As String
        Line Input #1, textline

        If nLine = 2 Then

            Dim lineArray_temp() As String
            lineArray_temp = Split(textline, " ")

            Dim j As Integer
            j = 0

            Dim i As Integer
            For i = 0 To UBound(lineArray_temp)
                If lineArray_temp(i) <> "" Then
                    j = j + 1
                End If
            Next

            ReDim lineArray(j)
            j = 0
            For i = 0 To UBound(lineArray_temp)
                If lineArray_temp(i) <> "" Then
                    lineArray(j) = lineArray_temp(i)
                    j = j + 1
                End If
            Next

            Dim DX As Double, Xmin  As Double, Xmax As Double, Xstart As Double
            DX = Val(lineArray(3))
            Xmin = Val(lineArray(0))
            Xmax = Val(lineArray(1))
            Xstart = Val(lineArray(2))

            Dim DY As Double, Ymin  As Double, Ymax As Double, Ystart As Double
            DY = Val(lineArray(7))
            Ymin = Val(lineArray(4))
            Ymax = Val(lineArray(5))
            Ystart = Val(lineArray(6))
        ElseIf nLine = 5 Then
            lineArray = Split(textline, "'")
            x_label(i_file) = lineArray(1)
            Call replaceGreekCharacters(x_label(i_file))

            Dim countLess As Boolean
            countLess = False

            Dim x_label_len As Double
            x_label_len = 0

            Dim iter As Integer
            For iter = 1 To Len(x_label(i_file))
                If Mid(x_label(i_file), iter, 1) = "{" Or Mid(x_label(i_file), iter, 1) = "^" Or Mid(x_label(i_file), iter, 1) = "_" Then
                    countLess = True
                ElseIf Mid(x_label(i_file), iter, 1) = "}" Or Mid(x_label(i_file), iter, 1) = "\" Then
                    countLess = False
                Else
                    If Mid(x_label(i_file), iter, 1) = " " Or Mid(x_label(i_file), iter, 1) = "/" Then
                            x_label_len = x_label_len + 0.5
                    ElseIf countLess = False Then
                        x_label_len = x_label_len + 1
                    Else
                        x_label_len = x_label_len + 0.4
                    End If
                End If
            Next

            'MsgBox (x_label_len)

         ElseIf nLine = 6 Then
            lineArray = Split(textline, "'")
            y_label(i_file) = lineArray(1)
            Call replaceGreekCharacters(y_label(i_file))
        End If

    Next nLine

    ' Lecture du nombre de lignes du titre
    Line Input #1, textline

    ' Lecture du titre
    Dim nLineTitle As Integer
    For nLineTitle = 1 To Int(textline)
        Line Input #1, textline
    Next

    ' Lecture des données numériques
    Dim nLineValue As Integer
    nLineValue = 1

    Dim i_row As Integer
    i_row = 1

    ' Lecture du fichier .des jusqu'à la dernière ligne
    Do Until EOF(1)

        Do While True ' On lit la série jusqu'à tomber sur une ligne de 8888.0

            Line Input #1, textline
            textline = Trim(textline)

            If nLineValue = 1 Then ' Lecture du style de dessin

                lineArray_temp = Split(textline, " ")

                j = 0
                For i = 0 To UBound(lineArray_temp)
                    If lineArray_temp(i) <> "" Then
                        j = j + 1
                    End If
                Next

                ReDim lineArray(j)
                j = 0
                For i = 0 To UBound(lineArray_temp)
                    If lineArray_temp(i) <> "" Then
                        lineArray(j) = lineArray_temp(i)
                        j = j + 1
                    End If
                Next

                Dim plotStyle As Integer, plotColor As Integer, plotSymbol As Integer
                plotStyle = lineArray(0)
                plotColor = lineArray(1)
                plotSymbol = lineArray(2)

                i_row = 0

            Else

                lineArray_temp = Split(textline, " ")

                j = 0
                For i = 0 To UBound(lineArray_temp)
                    If lineArray_temp(i) <> "" Then
                        j = j + 1
                    End If
                Next

                ReDim lineArray(j)
                j = 0
                For i = 0 To UBound(lineArray_temp)
                    If lineArray_temp(i) <> "" Then
                        lineArray(j) = lineArray_temp(i)
                        j = j + 1
                    End If
                Next

                If (Abs(CDbl(Val(Trim(lineArray(0)))) - 8888#) > 1E-16) And (Abs(CDbl(Val(Trim(lineArray(1)))) - 8888#) > 1E-16) Then
                    'récupération des valeurs "trimmées"
                    x_temp(i_row) = CDbl(Val(lineArray(0)))
                    y_temp(i_row) = CDbl(Val(lineArray(1)))

                    i_row = i_row + 1

                Else

                    nSeries = nSeries + 1
                    Set s = ActiveChart.SeriesCollection.NewSeries

                    Dim x() As Double
                    Dim y() As Double
                    ReDim x(i_row - 1)
                    ReDim y(i_row - 1)
                    For j = 0 To i_row - 1
                        x(j) = x_temp(j)
                        y(j) = y_temp(j)
                    Next

                    s.XValues = x
                    s.Values = y

                    If plotStyle = 0 Then
                        s.ChartType = xlXYScatterSmooth

                        Dim xlMarkerNone As Variant
                        s.MarkerStyle = xlMarkerNone
                        s.Format.Line.Visible = msoTrue

                        If plotColor = 1 Then
                            s.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
                        ElseIf plotColor = 2 Then
                            s.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                        ElseIf plotColor = 3 Then
                            s.Format.Line.ForeColor.RGB = RGB(0, 255, 0)
                        ElseIf plotColor = 4 Then
                            s.Format.Line.ForeColor.RGB = RGB(0, 255, 255)
                        ElseIf plotColor = 5 Then
                            s.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
                        ElseIf plotColor = 6 Then
                            s.Format.Line.ForeColor.RGB = RGB(255, 0, 255)
                        Else
                            s.Format.Line.ForeColor.RGB = RGB(0, 255, 255)
                        End If

                    ElseIf plotStyle = 1 Then
                        s.ChartType = xlXYScatter
                        s.Format.Line.Visible = msoTrue
                        s.Format.Line.Transparency = 1
                        s.MarkerSize = 3

                        If plotSymbol = 1 Then
                            s.MarkerStyle = xlMarkerStylePlus
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        ElseIf plotSymbol = 2 Then
                            s.MarkerStyle = xlMarkerStyleX
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        ElseIf plotSymbol = 3 Then
                            s.MarkerStyle = xlMarkerStyleSquare
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(255, 255, 255)
                        ElseIf plotSymbol = 4 Then
                            s.MarkerStyle = xlMarkerStyleDiamond
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(255, 255, 255)
                        ElseIf plotSymbol = 5 Then
                            s.MarkerStyle = xlMarkerStyleStar
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        ElseIf plotSymbol = 6 Then
                            s.MarkerStyle = xlMarkerStyleCircle
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(255, 255, 255)
                        ElseIf plotSymbol = 31 Then
                            s.MarkerStyle = xlMarkerStyleSquare
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        ElseIf plotSymbol = 41 Then
                            s.MarkerStyle = xlMarkerStyleDiamond
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        ElseIf plotSymbol = 61 Then
                            s.MarkerStyle = xlMarkerStyleSquare
                            s.MarkerForegroundColor = RGB(0, 0, 0)
                            s.MarkerBackgroundColor = RGB(0, 0, 0)
                        End If

                    End If

                    If plotColor = 1 Then
                        s.MarkerForegroundColor = RGB(0, 0, 0)
                    ElseIf plotColor = 2 Then
                        s.MarkerForegroundColor = RGB(255, 0, 0)
                    ElseIf plotColor = 3 Then
                        s.MarkerForegroundColor = RGB(0, 255, 0)
                    ElseIf plotColor = 4 Then
                        s.MarkerForegroundColor = RGB(0, 255, 0)
                    ElseIf plotColor = 5 Then
                        s.MarkerForegroundColor = RGB(0, 0, 255)
                    ElseIf plotColor = 6 Then
                        s.MarkerForegroundColor = RGB(255, 0, 255)
                    Else
                        s.MarkerForegroundColor = RGB(0, 255, 255)
                    End If

                    s.Select
                    Application.CommandBars("Format Object").Visible = False
                    With Selection.Format.Line
                        .Visible = msoTrue
                        .Weight = 0.75
                    End With

                    nLineValue = 1
                    i_row = 1
                    Erase x
                    Erase y

                    Exit Do
                End If

            End If

            nLineValue = nLineValue + 1

        Loop

    Loop

    Close #1

    ' Réglage de l'épaisseur des marques et ajout de la transparence aux marqueurs
    For i = 1 To nSeries
        ActiveChart.FullSeriesCollection(i).Format.Line.Weight = 0.3
        ActiveChart.FullSeriesCollection(i).Format.Fill.Visible = msoFalse
    Next

    ActiveChart.Axes(xlCategory).MinimumScale = Xmin
    ActiveChart.Axes(xlCategory).MaximumScale = Xmax
    ActiveChart.Axes(xlValue).MinimumScale = Ymin
    ActiveChart.Axes(xlValue).MaximumScale = Ymax

    ' Suppression de la légende
    ActiveChart.Legend.Delete

    ' Modification de la taille de l'objet graphique
    sh.Select
    sh.Height = 200
    sh.Width = 250

    ' Modification de la taille du graphique contenu dans l'objet précédemment redimensionné
    ActiveChart.PlotArea.Width = 200#

    ' Noms des axes
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = x_label(i_file)
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = y_label(i_file)
    End With

    ' Modification de la police et de sa taille
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Helvetica"
        .NameFarEast = "Helvetica"
        .Name = "Helvetica"
    End With

    Selection.Format.TextFrame2.TextRange.Font.Size = 6

    ' on déplace le titre des x
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Left = 196#
    Selection.Top = 165#

    ' déplacement et rotation du titre des y
    ActiveChart.Axes(xlValue).AxisTitle.Select
    Selection.Left = 202#
    Selection.Top = 17#
    Selection.Orientation = xlHorizontal
    Application.CommandBars("Format Object").Visible = False

    ' Création des portions droites de flèches
    ActiveChart.Shapes.AddConnector(msoConnectorStraight, 200, 177.7, _
        200 + x_label_len * 5, 177.7).Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoConnectorStraight
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.5
    End With

    ActiveChart.Shapes.AddConnector(msoConnectorStraight, 202, _
        30, 202, 13.1).Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoConnectorStraight
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.5
    End With

    ' création des pointes des flèches
    ActiveChart.Shapes.AddShape(msoShapeIsoscelesTriangle, 200 + x_label_len * 5 - 1.2, _
        176.5, 5, 2.5).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.5
    End With
    Selection.ShapeRange.Rotation = 90

    ActiveChart.Shapes.AddShape(msoShapeIsoscelesTriangle, 199.5, 8 + 2.3, 5, 2.5).Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Weight = 0.5
    End With

    ' Suppression des lignes horizontales dans la figure et ajout d'un cadre noir
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.PlotArea.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
    End With

    ' Abscisse mise en noir et épaisissement
    ActiveChart.Axes(xlCategory).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Application.CommandBars("Format Object").Visible = False
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
    End With

    ' Ordonnée mise en noir et épaisissement
    ActiveChart.Axes(xlValue).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Application.CommandBars("Format Object").Visible = False
    With Selection.Format.Line
        .Visible = msoTrue
        .Weight = 0.5
    End With


    ActiveChart.Axes(xlCategory).MajorUnit = 2 * DX
    ActiveChart.Axes(xlValue).MajorUnit = 2 * DY


    ' Création des ticks mineures de l'axe des abscisses
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinorUnit = ActiveChart.Axes(xlCategory).MajorUnit / 2
    Selection.MinorTickMark = xlOutside

    ' Création des ticks mineures de l'axe des ordonnées
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinorUnit = ActiveChart.Axes(xlValue).MajorUnit / 2
    Selection.MinorTickMark = xlOutside

    ' Les valeurs des tiques sont mises en gras et en italique
    ActiveChart.Axes(xlCategory).TickLabels.Font.Bold = msoTrue
    ActiveChart.Axes(xlValue).TickLabels.Font.Bold = msoTrue

    ActiveChart.Axes(xlValue).AxisTitle.Font.Italic = msoTrue
    ActiveChart.Axes(xlCategory).AxisTitle.Font.Italic = msoTrue

Next

ActiveSheet.Select

Dim WDapp As Variant

Set WDapp = CreateObject("Word.Application")
WDapp.Visible = False

Dim WDdoc As Variant

If FileLocked(WordFileName) Then

    ' Suppression la feuille "Graphics"
    Application.DisplayAlerts = False
    For iSheet = 1 To Sheets.Count
        If Sheets(iSheet).Name = "Graphics" Then
            Sheets(iSheet).Delete
        End If
    Next
    Application.DisplayAlerts = True

    ' Arrêt de l'exécution de la macro
    Exit Sub
End If

Set WDdoc = WDapp.Documents.Open(Filename:=WordFileName, ReadOnly:=False)

Dim subdoc() As Variant
ReDim subdoc(ActiveSheet.ChartObjects.Count)

Dim iCht As Integer
Dim PPT As Object
Dim newPres() As Object

ReDim newPres(ActiveSheet.ChartObjects.Count)

Dim PPSlide As Object

Set PPT = CreateObject("PowerPoint.Application")
PPT.Visible = True

For iCht = ActiveSheet.ChartObjects.Count To 1 Step -1

   ' copy objectchart as a picture
    ActiveSheet.ChartObjects(iCht).Copy

    ''''''''''''''
    ' POWERPOINT '
    ''''''''''''''

    Set newPres(iCht) = PPT.Presentations.Add(True)
    Set PPSlide = newPres(iCht).Slides.Add(1, 1)

    ' Suppression de tous les éléments de la slide (titre, zone de texte...)
'https://stackoverflow.com/questions/22811544/delete-all-shapes-of-a-powerpoint-slide
    PPSlide.Shapes.Range.Delete

    ' Modification de la taille de la slide
    With newPres(iCht).PageSetup
        .SlideWidth = 260
        .SlideHeight = 200
    End With

    ' Paste chart
    newPres(iCht).Slides(1).Shapes.Paste.Select

    ' on gère les indices et exposants ici car seul Powerpoint peut s'en occuper, pas Excel
   With newPres(iCht).Slides(1).Shapes(1)
        If (.HasChart = True) Then
            If (InStr(.Chart.Axes(xlCategory).AxisTitle.Characters.Text, "{") <> 0) Then
                With .Chart.Axes(xlCategory)

                    Dim mustSub As Boolean
                    mustSub = False
                    Dim mustSup As Boolean
                    mustSup = False
                    .AxisTitle.Characters.Text = "              "
                    Dim jChar As Integer
                    jChar = 1

                    Dim iChar As Integer
                    For iChar = 1 To Len(x_label(iCht))

                        If Mid(x_label(iCht), iChar, 1) = "^" Or Mid(x_label(iCht), iChar, 1) = "{" Then
                            If Mid(x_label(iCht), iChar + 1, 1) = "{" Then
                                mustSup = True
                            End If
                        ElseIf Mid(x_label(iCht), iChar, 1) = "_" Or Mid(x_label(iCht), iChar, 1) = "{" Then
                            If Mid(x_label(iCht), iChar + 1, 1) = "{" Then
                                mustSub = True
                            End If
                        ElseIf Mid(x_label(iCht), iChar, 1) = "}" Then
                            mustSub = False
                            mustSup = False
                        Else
                            .AxisTitle.Characters(jChar, 1).Text = Mid(x_label(iCht), iChar, 1)
                            If mustSub = True Then
                                .AxisTitle.Characters(Start:=jChar, Length:=1).Font.Subscript = True
                            ElseIf mustSup = True Then
                                .AxisTitle.Characters(Start:=jChar, Length:=1).Font.Superscript = True
                            End If
                            jChar = jChar + 1
                        End If
                    Next

                End With

            End If

        End If

        If (.HasChart = True) And (InStr(.Chart.Axes(xlValue).AxisTitle.Characters.Text, "{") <> 0) Then

            With .Chart.Axes(xlValue)

                mustSub = False
                mustSup = False
                .AxisTitle.Characters.Text = "              "
                jChar = 1

                For iChar = 1 To Len(y_label(iCht))

                    If Mid(y_label(iCht), iChar, 1) = "^" Or Mid(y_label(iCht), iChar, 1) = "{" Then
                        If Mid(y_label(iCht), iChar + 1, 1) = "{" Then
                            mustSup = True
                        End If
                    ElseIf Mid(y_label(iCht), iChar, 1) = "_" Or Mid(y_label(iCht), iChar, 1) = "{" Then
                        If Mid(y_label(iCht), iChar + 1, 1) = "{" Then
                            mustSub = True
                        End If
                    ElseIf Mid(y_label(iCht), iChar, 1) = "}" Then
                        mustSub = False
                        mustSup = False
                    Else
                        .AxisTitle.Characters(jChar, 1).Text = Mid(y_label(iCht), iChar, 1)
                        If mustSub = True Then
                            .AxisTitle.Characters(Start:=jChar, Length:=1).Font.Subscript = True
                        ElseIf mustSup = True Then
                            .AxisTitle.Characters(Start:=jChar, Length:=1).Font.Superscript = True
                        End If
                        jChar = jChar + 1
                    End If
                Next

            End With
        End If

    End With

    ' On enregistre la présentation contenant la slide
    newPres(iCht).SaveAs directory & "Plot" & Str(iCht) & ".pptx"
    newPres(iCht).Saved = True
    Dim WAIT As Double

    ' Ce timer permet de laisser le temps à Powerpoint de sauvegarder la présentation
    ' La variable timeToWait correspond au temps en secondes donné à Powerpoint pour sauvegarder
    ' Augmenter la valeur de timeToWait s'il y a un problème au moment où Powerpoint se ferme
    Dim timeToWait As Double
    timeToWait = 1
    WAIT = Timer
    While Timer < WAIT + timeToWait
       DoEvents  'do nothing
    Wend

    ''''''''
    ' WORD '
    ''''''''

    WDapp.Documents(WDdoc).Activate
    WDdoc.InlineShapes.AddOLEObject ClassType:="PowerPoint.Show.12", _
    Filename:=directory & "Plot" & Str(iCht) & ".pptx", LinkToFile:=False, _
    DisplayAsIcon:=False

    newPres(iCht).Close

    ' Suppression des fichiers Powerpoint créés
    ' https://stackoverflow.com/questions/67835/deleting-a-file-in-vba
    Dim filetodelete As String
    filetodelete = directory & "Plot" & Str(iCht) & ".pptx"
    ' First remove readonly attribute, if set
      SetAttr filetodelete, vbNormal
    ' Then delete the file
      Kill filetodelete

Next

' Fermeture de Powerpoint
PPT.Quit
Set PPT = Nothing

' Suppression des graphiques temporaires sur Excel
ActiveSheet.ChartObjects.Delete

Application.DisplayAlerts = True
Sheets("EXCELplorer").Activate
Application.ScreenUpdating = True

' Suppression la feuille "Graphics"
Application.DisplayAlerts = False
For iSheet = 1 To Sheets.Count
    If Sheets(iSheet).Name = "Graphics" Then
        Sheets(iSheet).Delete
    End If
Next

' on remet le délimiteur initial d'Excel
Application.DecimalSeparator = delim

' Affichage du docuent Word
WDapp.Documents.Open (WordFileName)

End Sub

Private Sub openDialog(txtFileName() As String, fileNumber As Integer)

    ' source : http://stackoverflow.com/questions/10304989/open-windows-explorer-and-select-a-file#10305150

    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd

      .AllowMultiSelect = True

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Fichier DESExplorer", "*.des"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      Dim i As Integer
      i = 0
      If .Show = True Then
        ReDim txtFileName(.SelectedItems.Count - 1)
        Dim Item As Variant
        For Each Item In .SelectedItems
            txtFileName(i) = Item 'replace txtFileName with your textbox
            i = i + 1
        Next
      End If
   End With

   fileNumber = i

End Sub

Private Sub selectWordFile(txtFileName As String)

    ' source : http://stackoverflow.com/questions/10304989/open-windows-explorer-and-select-a-file#10305150

    Dim fd As Office.FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

   With fd

      .AllowMultiSelect = False

      ' Set the title of the dialog box.
      .Title = "Please select the file."

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Fichier Word", "*.docx"
      .Filters.Add "All Files", "*.*"

      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        Dim Item As Variant
        For Each Item In .SelectedItems
            txtFileName = Item 'replace txtFileName with your textbox
        Next
      End If
   End With

End Sub

Sub replaceGreekCharacters(txt As String)

    txt = Replace(txt, "\Alpha", ChrW(&H391))
    txt = Replace(txt, "\Beta", ChrW(&H392))
    txt = Replace(txt, "\Gamma", ChrW(&H393))
    txt = Replace(txt, "\Delta", ChrW(&H394))
    txt = Replace(txt, "\Epsilon", ChrW(&H395))
    txt = Replace(txt, "\Zeta", ChrW(&H396))
    txt = Replace(txt, "\Eta", ChrW(&H397))
    txt = Replace(txt, "\Theta", ChrW(&H398))
    txt = Replace(txt, "\Iota", ChrW(&H399))
    txt = Replace(txt, "\Kappa", ChrW(&H39A))
    txt = Replace(txt, "\Lambda", ChrW(&H39B))
    txt = Replace(txt, "\Mu", ChrW(&H39C))
    txt = Replace(txt, "\Nu", ChrW(&H39D))
    txt = Replace(txt, "\Xi", ChrW(&H39E))
    txt = Replace(txt, "\Omicron", ChrW(&H39F))
    txt = Replace(txt, "\Pi", ChrW(&H3A0))
    txt = Replace(txt, "\Rho", ChrW(&H3A1))
    txt = Replace(txt, "\Sigma", ChrW(&H3A3))
    txt = Replace(txt, "\Tau", ChrW(&H3A4))
    txt = Replace(txt, "\Upsilon", ChrW(&H3A5))
    txt = Replace(txt, "\Phi", ChrW(&H3A6))
    txt = Replace(txt, "\Chi", ChrW(&H3A7))
    txt = Replace(txt, "\Psi", ChrW(&H3A8))
    txt = Replace(txt, "\Omega", ChrW(&H3A9))

    txt = Replace(txt, "\alpha", ChrW(&H3B1))
    txt = Replace(txt, "\beta", ChrW(&H3B2))
    txt = Replace(txt, "\gamma", ChrW(&H3B3))
    txt = Replace(txt, "\delta", ChrW(&H3B4))
    txt = Replace(txt, "\epsilon", ChrW(&H3B5))
    txt = Replace(txt, "\zeta", ChrW(&H3B6))
    txt = Replace(txt, "\eta", ChrW(&H3B7))
    txt = Replace(txt, "\theta", ChrW(&H3B8))
    txt = Replace(txt, "\iota", ChrW(&H3B9))
    txt = Replace(txt, "\kappa", ChrW(&H3BA))
    txt = Replace(txt, "\lambda", ChrW(&H3BB))
    txt = Replace(txt, "\mu", ChrW(&H3BC))
    txt = Replace(txt, "\nu", ChrW(&H3BD))
    txt = Replace(txt, "\xi", ChrW(&H3BE))
    txt = Replace(txt, "\omicron", ChrW(&H3BF))
    txt = Replace(txt, "\pi", ChrW(&H3C0))
    txt = Replace(txt, "\rho", ChrW(&H3C1))
    txt = Replace(txt, "\sigma", ChrW(&H3C3))
    txt = Replace(txt, "\tau", ChrW(&H3C4))
    txt = Replace(txt, "\upsilon", ChrW(&H3C5))
    txt = Replace(txt, "\phi", ChrW(&H3C6))
    txt = Replace(txt, "\chi", ChrW(&H3C7))
    txt = Replace(txt, "\psi", ChrW(&H3C8))
    txt = Replace(txt, "\omega", ChrW(&H3C9))


End Sub

Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next

   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Lock Read As #1
   Close #1

   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
      FileLocked = True
      Err.Clear
   End If
End Function
