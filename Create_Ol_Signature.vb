Sub CreateHTMLSignatures()
    ' Spécifier le chemin d'accès complet et le nom du fichier Excel contenant les informations des collaborateurs
        Dim ExcelFile As String
        ExcelFile = "Path\excelFile"
    
    ' Spécifier le chemin d'accès complet et le nom du modèle HTML
        Dim SignatureFileHTML As String
        SignatureFileHTML = "Path\templateHTML"
    
    ' Spécifier le chemin d'accès complet et le nom du modèle RTF
        Dim SignatureFileRTF As String
        SignatureFileRTF = "Path\templateRTF"
    
    ' Spécifier le chemin de sauvegarde des signatures
        Dim SignatureFolder As String
        SignatureFolder = "Path\saveFolder"
    
    ' Ouvrir le fichier Excel
        Dim xlApp As Object
        Dim xlWorkbook As Object
        Set xlApp = CreateObject("Excel.Application")
        Set xlWorkbook = xlApp.Workbooks.Open(ExcelFile)
    
    
    ' Parcourir toutes les lignes du fichier Excel et créer une signature pour chaque personne
        Dim iRow As Integer
        For iRow = 2 To xlWorkbook.Sheets(1).Cells(Rows.Count, 1).End(-4162).Row
        
    ' Récupérer les informations de la personne à partir de la ligne en cours
        Dim Nom As String
        Dim Prenom As String
        Dim Mail As String
        Dim Poste As String
        Dim Service As String
        Dim Num_fixe As String
        Dim Num_gsm As String
        Nom = xlWorkbook.Sheets(1).Cells(iRow, 1).Value
        Prenom = xlWorkbook.Sheets(1).Cells(iRow, 2).Value
        Mail = xlWorkbook.Sheets(1).Cells(iRow, 3).Value
        Poste = xlWorkbook.Sheets(1).Cells(iRow, 4).Value
        Service = xlWorkbook.Sheets(1).Cells(iRow, 5).Value
        Num_gsm = xlWorkbook.Sheets(1).Cells(iRow, 6).Value
        Num_fixe = xlWorkbook.Sheets(1).Cells(iRow, 7).Value


        
    ' Créer une nouvelle signature HTML pour la personne en cours
        Dim oSig As Outlook.MailItem
        Set oSig = CreateObject("Outlook.Application").CreateItem(olMailItem)
        
    ' Charger le contenu du fichier HTML dans la signature
        Dim SigTextHTML As String
        Open SignatureFileHTML For Input As #1
        SigTextHTML = Input$(LOF(1), #1)
        Close #1
        SigTextHTML = Replace(SigTextHTML, "Nom", Nom)
        SigTextHTML = Replace(SigTextHTML, "Prenom", Prenom)
        SigTextHTML = Replace(SigTextHTML, "Mail", Mail)
        SigTextHTML = Replace(SigTextHTML, "Poste", Poste)
        SigTextHTML = Replace(SigTextHTML, "Service", Service)
        SigTextHTML = Replace(SigTextHTML, "Num_fixe", Num_fixe)
        SigTextHTML = Replace(SigTextHTML, "Num_gsm", Num_gsm)
        oSig.HTMLBody = SigTextHTML
        
        
    ' Enregistrer la signature dans le dossier spécifié
        Dim SignatureNameHTML As String
        SignatureNameHTML = Nom & "_" & Prenom & "_" & "Sig" & ".htm"
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim SignaturePathHTML As String
        SignaturePathHTML = SignatureFolder & "\" & SignatureNameHTML
        Dim tsHTML As Object
        Set tsHTML = fso.CreateTextFile(SignaturePathHTML, True)
        tsHTML.Write oSig.HTMLBody
        tsHTML.Close
        
    ' Charger le contenu du fichier RTF dans la signature
        Dim SigTextRTF As String
        Open SignatureFileRTF For Input As #1
        SigTextRTF = Input$(LOF(1), #1)
        Close #1
        SigTextRTF = Replace(SigTextRTF, "Nom", Nom)
        SigTextRTF = Replace(SigTextRTF, "Prenom", Prenom)
        SigTextRTF = Replace(SigTextRTF, "Mail", Mail)
        SigTextRTF = Replace(SigTextRTF, "Poste", Poste)
        SigTextRTF = Replace(SigTextRTF, "Service", Service)
        SigTextRTF = Replace(SigTextRTF, "tel", Num_fixe)
        SigTextRTF = Replace(SigTextRTF, "gsm", Num_gsm)
        oSig.Body = SigTextRTF
        
    ' Enregistrer la signature dans le dossier spécifié
        Dim SignatureNameRTF As String
        SignatureNameRTF = Nom & "_" & Prenom & "_" & "Sig" & ".rtf"
        Dim fsoRTF As Object
        Set fsoRTF = CreateObject("Scripting.FileSystemObject")
        Dim SignaturePathRTF As String
        SignaturePathRTF = SignatureFolder & "\" & SignatureNameRTF
        Dim tsRTF As Object
        Set tsRTF = fsoRTF.CreateTextFile(SignaturePathRTF, True)
        tsRTF.Write oSig.Body
        tsRTF.Close
        
    'Enregistrer la signature au format TXT
        Dim SignatureNameTXT As String
        SignatureNameTXT = Nom & "_" & Prenom & "_" & "Sig" & ".txt"
        Dim SignaturePathTXT As String
        SignaturePathTXT = SignatureFolder & "\" & SignatureNameTXT
        Dim TXTFileOutput As Object
        Set TXTFileOutput = CreateObject("Word.Application")
        Dim Doc As Object
        Set Doc = TXTFileOutput.Documents.Open(SignaturePathRTF, False, True)
        Doc.SaveAs2 Filename:=SignaturePathTXT, FileFormat:=2
        Doc.Close False
        TXTFileOutput.Quit

    ' Supprimer la signature de la collection des signatures dans Outlook
        oSig.Delete
    
        Next iRow
    
    ' Fermer le fichier Excel
        xlWorkbook.Close SaveChanges:=False
        Set xlWorkbook = Nothing
        Set xlApp = Nothing