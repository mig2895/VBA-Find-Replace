'Le logiciel a été développé par Miguel de Bruyne dans le cadre du projet FNS SPAR, Open Education for Research Methodology Teaching across the Mediterranean, http://p3.snf.ch/project-190634

'Il est possible qu'il faut ajouter les références :
' - Microsoft Word 16.0
' - Microsoft VBScript Regular Expressions
' - Microsoft Scripting Runtime
'sous "outils" puis "Références"
Option Compare Text
Public combinedArr As Variant

Sub Starting()
    
    'Création d'un tableau à partir des colonnes A et B.
    'Le logiciel sait exactement quand arreter la recherche.
    
    With Sheets("Sheet1")
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    For i = 2 To LastRow
        .Cells(i, "A").Value = Trim(.Cells(i, "A").Value)
    Next i
    End With
    'Debug.Print LastRow
    
    With Application
    aArr = .Transpose(Range("A3:A" & LastRow).Value)
    bArr = .Transpose(Range("B3:B" & LastRow).Value)
    combinedArr = .Transpose(Array(aArr, bArr))
    End With

    'Debug.Print combinedArr(1, 1)
    
    
    'Cette fonction permet de lancer l'application
    Dim fileName As Variant
    Dim Path As String
    Dim PathFromFunction As String
    PathFromFunction = GetFolder()
    Path = PathFromFunction & "\"
    
    Call LoopAllSubFolders(Path)
    
    
    MsgBox ("Fini")
End Sub


Sub LoopAllSubFolders(ByVal folderPath As String)
    'Cette fonction va chercher tous les documents sous le dossier indiqué. Lors qu'il a établie la liste il va
    'La parcourir et lancer la fonction docSearch() pour chaque document.
    
    Dim fileName As String
    Dim fullFilePath As String
    Dim numFolders As Long
    Dim folders() As String
    Dim i As Long
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = Dir(folderPath & "*.*", vbDirectory)
    
    While Len(fileName) <> 0
    
        If Left(fileName, 1) <> "." Then
     
            fullFilePath = folderPath & fileName
     
            If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
                ReDim Preserve folders(0 To numFolders) As String
                folders(numFolders) = fullFilePath
                numFolders = numFolders + 1
            Else
                'Appelle la fonction docsearch avec le chemin du document.
                'Vérification que le fichier soit bien un document word.
                Debug.Print fullFilePath
                If InStr(1, fullFilePath, ".docx", vbTextCompare) > 0 Then
                    Call docSearch(fullFilePath)
                    Call CloseWordDocuments
                End If
            End If
     
        End If
     
        fileName = Dir()
    
    Wend
    
    For i = 0 To numFolders - 1
    
        LoopAllSubFolders folders(i)
     
    Next i
    
    

End Sub

Sub docSearch(ByVal FullDocumentPath As String)
    'Cette fonction va permettre d'ouvrir un document word, chercher les mots qui sont sur la liste des mots
    ' à chercher, remplacera les mots par l'autre liste, sauvegardera le fichier, puis le fermera.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdRng As Word.Range
    
    
    With Sheets("Sheet1")
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    End With
    
    
    Set wdApp = CreateObject("word.application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(FullDocumentPath)
    For Each wdRng In wdDoc.StoryRanges
    

        'Remplacement des mots
        Dim i As Integer
        For i = LBound(combinedArr) To UBound(combinedArr)
        If InStr(wdRng, combinedArr(i, 1)) Then
            With wdRng.Find
                    .Text = combinedArr(i, 1)
                    .Replacement.Text = "[" & combinedArr(i, 2) & "]"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchWholeWord = True
                    .Execute replace:=wdReplaceAll
                End With
        End If
        Next i
                
        'Sauvegarde du document
        wdDoc.Save
    
    
    Next wdRng
    
    'Fermeture de l'application Word
    wdApp.Quit
    Set wdApp = Nothing: Set wdDoc = Nothing: Set wdRng = Nothing
    
    
End Sub


Sub CloseWordDocuments()

    Dim objWord As Object

    Do
        On Error Resume Next
        Set objWord = GetObject(, "Word.Application")
        If Not objWord Is Nothing Then
            objWord.Quit
            Set objWord = Nothing
        End If
    Loop Until objWord Is Nothing

End Sub

Sub ListLoop()
      ' Select cell A2, *first line of data*.
      Range("A1").Select
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
         ' Insert your code here.
         ' Step down 1 row from present location.
         ActiveCell.Offset(1, 0).Select
      Loop
   End Sub
   
   Function GetFolder() As String
   'Cette fonction s'occupe de demander à l'utilisateur de choisir le dossier qu'il veut pseudoomiser
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Veuillez sélectionner le dossier contenant les fichiers à pseudomiser"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function


Function lrow(ByVal colNum As Long) As Long
    'Cette fonction permet de trouver la dernière ligne du document exel
    'Elle permet également d'enlever les espaces inutiles avant et après le texte
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lrow
        .Cells(i, "A").Value = Trim(.Cells(i, "A").Value)
    Next i
End Function



