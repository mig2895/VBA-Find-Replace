'Le logiciel a été développé par Miguel de Bruyne dans le cadre du projet FNS SPAR, Open Education for Research Methodology Teaching across the Mediterranean, http://p3.snf.ch/project-190634
Option Compare Text
Public combinedArr As Variant
 Global j As Integer
 
 Sub Start()

 
 j = 1
 
 'Cette fonction permet de lancer l'application
    Dim fileName As Variant
    Dim Path As String
    Dim PathFromFunction As String
    PathFromFunction = GetFolder()
    Path = PathFromFunction & "\"
    
    Call LoopAllSubFolders(Path)
    
    
    MsgBox ("Fini")
 
 End Sub
 
 
 
 
 Function GetFolder() As String
   'Cette fonction s'occupe de demander à l'utilisateur de choisir le dossier qu'il veut pseudoomiser
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Veuillez sélectionner le dossier contenant les fichiers à documenter"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

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
    Call HyperlinkFileListing(folderPath)
    
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
                
                
                
                
            End If
     
        End If
     
        fileName = Dir()
    
    Wend
    
    For i = 0 To numFolders - 1
    
        LoopAllSubFolders folders(i)
     
    Next i
    
    

End Sub

Sub HyperlinkFileListing(fullFilePath)
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = fullFilePath
'Get the folder object
Set objFolder = objFSO.GetFolder(objStartFolder)

'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
    'select cell
    Range(Cells(j + 1, 1), Cells(j + 1, 1)).Select
    'create hyperlink in selected cell
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        objFile.Path, _
        TextToDisplay:=objFile.Name
    j = j + 1
Next objFile
End Sub
