Attribute VB_Name = "Importation"
Option Explicit
Option Base 1 'Les tableaux commencent à l'index 1
Sub importation_prix_30stocks()
    'Procédure qui importe notre CSV puis le format dans une feuille de calcul nommé 'Cours'

    'Déclaration des classeurs
    Dim wbSource As Workbook
    
    'Déclaration des feuilles de calcul
    Dim wsPrix30Stocks As Worksheet
    
    'Déclaration des autres varaiables
    Dim path As Variant
    
    'Référence au classeur où se trouve ce code
    Set wbSource = ThisWorkbook
    
    'Boîte de dialogue pour choisir le fichier CSV
    path = Application.GetOpenFilename()
    
    'Vérifie si l'utilisateur a annulé
    If path = False Then
        MsgBox "Aucun fichier sélectionné"
        Exit Sub
    End If
    
    'Supprime la feuille "Prix 30 Stocks " si elle existe déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    wbSource.Sheets("Prix 30 Stocks").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Crée une nouvelle feuille "Prix 30 Stocks"
    Set wsPrix30Stocks = wbSource.Sheets.Add
    wsPrix30Stocks.Name = "Prix 30 Stocks"
    
    'Import du CSV dans la feuille Cours
    With wsPrix30Stocks.QueryTables.Add(Connection:="TEXT;" & path, Destination:=wsPrix30Stocks.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True   ' CSV avec virgule
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
    End With
    
    MsgBox "Le CSV a bien été importé dans la feuille 'Prix 30 Stocks'."
    
     'Mise en page du tableau "Prix 30 Stocks"
    With wsPrix30Stocks.Cells(1, 1).Resize(1, 31)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
    With wsPrix30Stocks.Cells(2, 1).Resize(121, 1)
        .Font.Bold = True
        .Interior.Color = RGB(164, 188, 43)
    End With
    
End Sub
Sub importation_prix_benchmark()
    'Procédure qui importe notre CSV puis le format dans une feuille de calcul nommé 'Benchmark'

    'Déclaration des classeurs
    Dim wbSource As Workbook
    
    'Déclaration des feuilles de calcul
    Dim wsPrixBench As Worksheet
    
    'Déclaration des autres varaiables
    Dim path As Variant
    
    'Référence au classeur où se trouve ce code
    Set wbSource = ThisWorkbook
    
    'Boîte de dialogue pour choisir le fichier CSV
    path = Application.GetOpenFilename()
    
    'Vérifie si l'utilisateur a annulé
    If path = False Then
        MsgBox "Aucun fichier sélectionné"
        Exit Sub
    End If
    
    'Supprime la feuille "Prix Bench" si elle existe déjà
    On Error Resume Next
    Application.DisplayAlerts = False
    wbSource.Sheets("Prix Bench").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    'Crée une nouvelle feuille "Prix Bench"
    Set wsPrixBench = wbSource.Sheets.Add
    wsPrixBench.Name = "Prix Bench"
    
    'Import du CSV dans la feuille Prix Bench
    With wsPrixBench.QueryTables.Add(Connection:="TEXT;" & path, Destination:=wsPrixBench.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True   ' CSV avec virgule
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
    End With
    
    MsgBox "Le CSV a bien été importé dans la feuille 'Prix Bench'."
    
    'Mise en page du tableau "Prix Bench"
    With wsPrixBench.Cells(1, 1).Resize(1, 2)
        .Font.Bold = True
        .Interior.Color = RGB(224, 224, 224)
    End With
     With wsPrixBench.Cells(2, 1).Resize(121, 1)
        .Font.Bold = True
        .Interior.Color = RGB(164, 188, 43)
    End With
    
End Sub
