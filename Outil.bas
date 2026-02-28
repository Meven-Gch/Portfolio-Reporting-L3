Attribute VB_Name = "Outil"
Option Explicit
Option Base 1
Sub tool_kit()
    'Procédure permettant de :
    
    ' - Calculer les rendements de nos 30 actions à partir de la feuille "Cours".
    
    ' - Les afficher dans une autres feuille nommé "Rendements".
    
    ' - Calculer les rendements du benchmark à partir de la feuille "Benchmark".
    
    ' - Les afficher dans une autres feuille nommé "Rend Bench".
    
    ' - Calculer les indicateurs clés de performance (Moyenne des Rendements, Moyenne Annualisé,
    'Ratio de Sharpe).
    
    '- Calculer les indicateurs clés de risques (Volatilité,Volatilité annualisé, VaR, TrackingError
    'Ratio d'Information).
    
    'ATTENTION : les données sont supposées être  mensuelles
    
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '% Déclaration des variables
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    'Déclaration des classeurs
    Dim wbSource As Workbook
    
    'Déclaration des feuilles de calcul
    Dim wsPrix30Stocks As Worksheet
    Dim wsRend30Stocks As Worksheet
    Dim wsPrixBench As Worksheet
    Dim wsRendBench As Worksheet
    Dim wsStats As Worksheet
    
    
    'Déclaration des variables itératives
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim c As Range
    
    'Déclaration des autres variables
    Dim lastRowR As Long
    Dim lastColR As Long
    Dim lastRowBench As Long
    Dim lastColBench As Long
    Dim rf As Double
    Dim tmp() As Double
    Dim TE As Double
    Dim meanR As Double, volR As Double
    Dim meanExcess As Double
    
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '% Mise en place des fichiers ("Rend 30 Stocks")
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    
    Set wbSource = ThisWorkbook
    Set wsPrix30Stocks = wbSource.Sheets("Prix 30 Stocks")

    'Supprimer la feuille "Rend 30 Stocks" si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    wbSource.Sheets("Rend 30 Stocks").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Création de la feuille "Rend 30 Stocks"
    Set wsRend30Stocks = wbSource.Sheets.Add
    wsRend30Stocks.Name = "Rend 30 Stocks"
   
    
    'Déterminer les dimensions de nos données
    lastRowR = wsPrix30Stocks.Cells(wsPrix30Stocks.Rows.Count, 2).End(xlUp).Row
    lastColR = wsPrix30Stocks.Cells(2, wsPrix30Stocks.Columns.Count).End(xlToLeft).Column
 
    'Forcer conversion des prix en nombres
    For Each c In wsPrix30Stocks.Range(wsPrix30Stocks.Cells(2, 2), wsPrix30Stocks.Cells(lastRowR, lastColR))
        If c.Value <> "" Then
            ' remplace la virgule par le séparateur décimal du système
            c.Value = CDbl(Replace(c.Value, ".", Application.DecimalSeparator))
        End If
    Next c
    
    'Copier les en-têtes
    wsRend30Stocks.Cells(1, 1).Value = "Date"
    wsPrix30Stocks.Range(wsPrix30Stocks.Cells(1, 2), wsPrix30Stocks.Cells(1, lastColR)).Copy wsRend30Stocks.Cells(1, 2)
    
    'Calcul des rendements pour chacune de nos actions sur toute la période
    For i = 3 To lastRowR
        wsRend30Stocks.Cells(i - 1, 1).Value = wsPrix30Stocks.Cells(i, 1).Value  ' Date
        
        For j = 2 To lastColR
            If wsPrix30Stocks.Cells(i - 1, j).Value <> 0 Then
                wsRend30Stocks.Cells(i - 1, j).Value = wsPrix30Stocks.Cells(i, j).Value / wsPrix30Stocks.Cells(i - 1, j).Value - 1
            End If
        Next j
    Next i
    
    'Formater en pourcentage
    wsRend30Stocks.Range(wsRend30Stocks.Cells(2, 2), wsRend30Stocks.Cells(lastRowR - 1, lastColR)).NumberFormat = "0.00%"
    
    'Mise en page du tableau "Rend 30 Stocks"
    With wsRend30Stocks.Cells(1, 1).Resize(1, 31)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
     With wsRend30Stocks.Cells(2, 1).Resize(121, 1)
        .Font.Bold = True
        .Interior.Color = RGB(164, 188, 43)
    End With
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '% Mise en place des fichiers ("Rend Bench")
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    
    Set wsPrixBench = wbSource.Sheets("Prix Bench")
    
    'Supprimer la feuille "Rend Bench" si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    wbSource.Sheets("Rend Bench").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Créationde la feuille "Rend Bench"
    Set wsRendBench = wbSource.Sheets.Add
    wsRendBench.Name = "Rend Bench"
    
     'Déterminer les dimensions de nos données
    lastRowBench = wsPrixBench.Cells(wsPrixBench.Rows.Count, 2).End(xlUp).Row
    lastColBench = wsPrixBench.Cells(2, wsPrixBench.Columns.Count).End(xlToLeft).Column
    
     'Forcer conversion des prix en nombres
    For Each c In wsPrixBench.Range(wsPrixBench.Cells(2, 2), wsPrixBench.Cells(lastRowBench, 2))
        If c.Value <> "" Then
            ' remplace la virgule par le séparateur décimal du système
            c.Value = CDbl(Replace(c.Value, ".", Application.DecimalSeparator))
        End If
    Next c
    
    'Copier les en-têtes
    wsRendBench.Cells(1, 1).Value = "Date"
    wsPrixBench.Range(wsPrixBench.Cells(1, 2), wsPrixBench.Cells(1, lastRowBench)).Copy wsRendBench.Cells(1, 2)
    
    'Calcul des rendements pour chacune de nos actions sur toute la période
    For i = 3 To lastRowBench
        wsRendBench.Cells(i - 1, 1).Value = wsPrixBench.Cells(i, 1).Value  ' Date
        
        If wsPrixBench.Cells(i - 1, 2).Value <> 0 Then
            wsRendBench.Cells(i - 1, 2).Value = wsPrixBench.Cells(i, 2).Value / wsPrixBench.Cells(i - 1, 2).Value - 1
        End If
    Next i
    
    'Formater en pourcentage
    wsRendBench.Range(wsRendBench.Cells(2, 2), wsRendBench.Cells(lastRowBench - 1, 2)).NumberFormat = "0.00%"
    
    'Mise en page du tableau "Rend Bench"
    With wsRendBench.Cells(1, 1).Resize(1, 31)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
     With wsRendBench.Cells(2, 1).Resize(121, 1)
        .Font.Bold = True
        .Interior.Color = RGB(164, 188, 43)
    End With
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '% Calcul des performances et des risques
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    'Choix du taux sans risque (0,3% par défaut)
    rf = InputBox("Entrer un taux sans risque :", , 0.003)
    
    MsgBox "Le taux sans risque est de :" & rf
    
    'Supprimer la feuille "Stats" si elle existe
    On Error Resume Next
    Application.DisplayAlerts = False
    wbSource.Sheets("Stats").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Créationde la feuille "Stats"
    Set wsStats = wbSource.Sheets.Add
    wsStats.Name = "Stats"
    
    'Calcul des indicateurs de performance et de risque
    wsStats.Cells(1, 1).Resize(1, 9).Value = Array("         Actifs         ", "Moy Rendements", "Moyenne Ann", "  Volatilté  ", "Volatilité Ann", "Sharpe Ratio", "Value at Risk (VaR)", " Tracking Error ", "Ratio d'Information")
    
    'Centrage par défaut des contenus des cellules
    wsStats.Cells.HorizontalAlignment = xlCenter
    wsStats.Cells.VerticalAlignment = xlCenter
    
    'Ajustement de la largeur des colonnes
    wsStats.UsedRange.Columns.AutoFit 'UsedRange pour toutes les cellules non-vide de stats
    wsStats.UsedRange.Rows.AutoFit
     
    For j = 2 To lastColR
    'Nom des actifs
    wsStats.Cells(j, 1).Value = wsRend30Stocks.Cells(1, j).Value
    
    'Rendement moyen
    meanR = Application.WorksheetFunction.Average(wsRend30Stocks.Range(wsRend30Stocks.Cells(2, j), wsRend30Stocks.Cells(lastRowR - 1, j)))
    wsStats.Cells(j, 2).Value = meanR
    wsStats.Cells(j, 4).Value = meanR * 12 ' annualisation
    
    'Volatilité
    volR = Application.WorksheetFunction.StDev(wsRend30Stocks.Range(wsRend30Stocks.Cells(2, j), wsRend30Stocks.Cells(lastRowR - 1, j)))
    wsStats.Cells(j, 3).Value = volR
    wsStats.Cells(j, 5).Value = volR * Sqr(12) ' annualisé
    
    'Sharpe
    If volR <> 0 Then
        wsStats.Cells(j, 6).Value = (meanR - rf) / volR * Sqr(12)
    End If
    
    'VaR 5%
    wsStats.Cells(j, 7).Value = WorksheetFunction.Norm_Inv(0.05, meanR, volR)
    
    'Tracking Error et Information Ratio
    ReDim tmp(1 To lastRowR - 1)
    For k = 2 To lastRowR - 1
        tmp(k - 1) = wsRend30Stocks.Cells(k, j).Value - wsRendBench.Cells(k, 2).Value
    Next k
    
    'TE
    TE = Application.WorksheetFunction.StDev(tmp)
    wsStats.Cells(j, 8).Value = TE * Sqr(12) ' annualisé
    
    'Information Ratio
    meanExcess = Application.WorksheetFunction.Average(tmp)
    If TE <> 0 Then
        wsStats.Cells(j, 9).Value = meanExcess / TE * Sqr(12)
    End If
    
Next j
   
    MsgBox "Indicateurs statisqtiques calculés avec succès !"

    'Titre des lignes de Total
    wsStats.Cells(34, 1).Value = "Total"
    
    'Calcul des totaux du tableau de la feuille Stat
    For j = 2 To 9
        wsStats.Cells(34, j).Value = Application.WorksheetFunction.Average(wsStats.Cells(2, j).Resize(30, 1))
    Next j
    
    'Mise en page des lignes Total
    With wsStats.Cells(34, 1)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
    With wsStats.Cells(34, 1).Resize(1, 10)
        .Font.Bold = True
    End With
    
    'Mise en page du tableau des statistiques
    With wsStats.Cells(1, 1).Resize(1, 9)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
     With wsStats.Cells(2, 1).Resize(30, 1)
        .Font.Bold = True
        .Interior.Color = RGB(147, 187, 243)
    End With

    
End Sub
