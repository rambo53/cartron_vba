' =========================================================================
' 1. PROCÉDURE PRINCIPALE (LIÉE AU BOUTON)
' =========================================================================
Sub Bouton1_Cliquer()
    ' ---------------------
    ' FORMULAIRE UserForm1
    ' ---------------------
    Dim FrmDates As New UserForm1
    FrmDates.Show ' Le code s'arrête ici tant que l'utilisateur n'a pas cliqué sur Valider
    
    ' Si l'utilisateur a cliqué sur la croix pour fermer
    If FrmDates.Annule Then
        Unload FrmDates
        Exit Sub ' On arrête proprement la macro
    End If
    
    ' On récupère les dates saisies dans des variables si tu en as besoin plus tard
    Dim DateDebutChoisie As Date
    Dim DateFinChoisie As Date
    DateDebutChoisie = FrmDates.DateDebut
    DateFinChoisie = FrmDates.DateFin
    
    ' On décharge définitivement le formulaire de la mémoire
    Unload FrmDates
    ' ---------------------
    ' FORMULAIRE UserForm1
    ' ---------------------
    
    
    ' ---------------------
    '   TRAITEMENT
    ' ---------------------
    
    ' --- CONSTANTES ---
    Const directory_in_name As String = "fichier_a_traiter"
    Const directory_out_name As String = "fichier_genere"
    Const directory_out_archive As String = "fichier_traite"
    Const Journal_export_achat As String = "Journal"
    Const Compte_export_achat As String = "Compte"
    Const Piece_export_achat As String = "Pièce"
    Const export_achat_libelle As String = "Libellé"
    Const Debit_export_achat As String = "Débit"
    Const Credit_export_achat As String = "Crédit"
    Const Col_defaut_coala As String = "Defaut"
    Const libelle_clean As String = "libelle_clean"
    Const key_table As String = "key_table"
    Const date_ech As String = "date_ech"
    Const IBAN As String = "IBAN"
    Const RIB As String = "RIB"
    Const BIC As String = "BIC"
    Const xml_name As String = "bank_file.xml"
    Const date_libelle As String = "Date"
    ' --- CONSTANTES ---
    
    Dim new_columns As Variant
    Dim export_achat As Worksheet
    Dim export_achat_name As String
    Dim table_correspondance As Worksheet
    Dim table_correspondance_name As String
    Dim Plage_libelle As Range
    Dim Plage_debit As Range
    Dim row_to_start As Long
    Dim col_to_inject_key_generate As Long
    Dim col_to_inject_key_find As Long
    Dim col_to_inject_iban As Long
    Dim col_to_inject_rib As Long
    Dim col_to_inject_bic As Long
    Dim PlageRecherche As Range
    Dim Col_correspondance As String
    Dim Cellule_libelle As Range
    Dim Cellule_correspondance As Range
    Dim CheminDossierIn As String
    Dim CheminDossierPrincipal As String
    Dim ClasseurPrincipal As Workbook
    Dim Cellule_libelle_clean As Range
    Dim Cellule_key_table As Range
    Dim Cellule_IBAN As Range
    Dim Cellule_RIB As Range
    Dim Cellule_BIC As Range
    Dim Cellule_date_ech As Range
    Dim Cellule_debit As Range
    Dim CheminDossierOut As String
    Dim CheminDossierArchive As String
    Dim number_of_transactions As Long
    Dim total_payments As Double
    Dim data_dict As Object
    Dim paiement_dict As Object
    
    new_columns = Array(date_libelle, Journal_export_achat, Compte_export_achat, Piece_export_achat, export_achat_libelle, _
                        Debit_export_achat, Credit_export_achat, Col_defaut_coala, libelle_clean, key_table, date_ech, IBAN, RIB, BIC)
    
    row_to_start = 2
    Col_correspondance = "correspondance_coala"
    table_correspondance_name = "table_correspondance"
    export_achat_name = "export_achat"
    
    Set ClasseurPrincipal = ThisWorkbook
    CheminDossierPrincipal = ClasseurPrincipal.Path
    
    ' récupération et vérification des chemins pour les fichiers de données
    CheminDossierIn = get_path_directory(CheminDossierPrincipal, directory_in_name)
    CheminDossierOut = get_path_directory(CheminDossierPrincipal, directory_out_name)
    CheminDossierArchive = get_path_directory(CheminDossierPrincipal, directory_out_archive)
    
    inject_data_in_worbook CheminDossierIn, export_achat_name, ClasseurPrincipal
    
    ' --- APPEL DES FONCTIONS POUR DÉFINIR LA PLAGE ---
    Set export_achat = Get_sheet(export_achat_name)
    
    ' SUB création des nouvelles colonnes dans l'export_achat
    create_columns new_columns, export_achat
    
    ' 1. Appel de la fonction letter_colonne pour obtenir la lettre de la cellule d'un nom de colonne
    Set Cellule_libelle = letter_colonne(export_achat_libelle, export_achat)
   
    ' 2. Appel de la fonction Plage_to_check pour générer la plage de données (Objets = SET)
    Set Plage_libelle = Plage_to_check(Cellule_libelle, export_achat, export_achat_name, row_to_start)
    
    ' lecture de la table de l'onglet "table_correspondance"
    Set table_correspondance = Get_sheet(table_correspondance_name)
    
    '
    Set Cellule_correspondance = letter_colonne(Col_correspondance, table_correspondance)
    ' Définir la zone de recherche dans la table de correspondance
    Set PlageRecherche = Plage_to_check(Cellule_correspondance, table_correspondance, table_correspondance_name, row_to_start)
    
    Set Cellule_libelle_clean = letter_colonne(libelle_clean, export_achat)
    Set Cellule_key_table = letter_colonne(key_table, export_achat)
    Set Cellule_IBAN = letter_colonne(IBAN, export_achat)
    Set Cellule_RIB = letter_colonne(RIB, export_achat)
    Set Cellule_BIC = letter_colonne(BIC, export_achat)
    Set Cellule_date_ech = letter_colonne(date_ech, export_achat)
    
    'SUB génération de la matrice complétée
    generate_matrice Plage_libelle, PlageRecherche, Cellule_libelle_clean, Cellule_key_table, Cellule_IBAN, Cellule_RIB, Cellule_BIC, Cellule_date_ech, IBAN, RIB, BIC
    
    ' on récupère les informations de position de la colonne "Débit"
    Set Cellule_debit = letter_colonne(Debit_export_achat, export_achat)
    Set Plage_debit = Plage_to_check(Cellule_debit, export_achat, export_achat_name, row_to_start)
    
    'SUB traitement des débits
    get_debit Plage_debit, PlageRecherche, Cellule_libelle, Cellule_IBAN, Cellule_RIB, Cellule_BIC, Cellule_date_ech, Cellule_key_table, IBAN, RIB, BIC

    
    'je récupère mon dictionnaire issue de l'onglet "DATA"
    Set data_dict = get_data_dict()
    
    ' SUB génération du Xml
    generate_xml CheminDossierOut, xml_name, export_achat, Cellule_date_ech, Cellule_libelle_clean, row_to_start, Credit_export_achat, Debit_export_achat, BIC, key_table, IBAN, RIB, date_ech, number_of_transactions, total_payments, paiement_dict, DateDebutChoisie, DateFinChoisie, Compte_export_achat, data_dict
    
    ' SUB archivage du fichier traité
    archive_file CheminDossierArchive, CheminDossierIn, CheminDossierOut, xml_name
    
    Dim Message As String
    Dim Message_dict As String
    
    For Each debit In paiement_dict.Keys
        Message_dict = Message_dict & paiement_dict(debit) & vbCrLf
    Next debit
    
    Message = "Fichier XML validé et généré avec succès !" & vbCrLf & vbCrLf & _
              "Résumé du traitement :" & vbCrLf & _
              "• Nombre de transactions : " & number_of_transactions & vbCrLf & _
              "• Montant total cumulé : " & Format(total_payments, "#,##0.00 €") & vbCrLf & _
              "Détails :" & vbCrLf & _
              Message_dict
    MsgBox Message, vbInformation, "Génération SEPA Terminée"
    
    ' ---------------------
    '   TRAITEMENT
    ' ---------------------
    
End Sub


' =========================================================================
' 2. LES FONCTIONS EXTÉRIEURES (INDÉPENDANTES)
' =========================================================================

Function letter_colonne(ByVal label_to_find As String, ByVal sheet_to_use As Worksheet) As Range
    Dim CelluleTrouvee As Range
    ' On renvoie un Objet Range (la cellule du titre), donc on utilise "As Range" et le mot "Set"
    Set CelluleTrouvee = sheet_to_use.Rows(1).Find(What:=label_to_find, LookIn:=xlValues, LookAt:=xlWhole)
    If CelluleTrouvee Is Nothing Then
        MsgBox "Erreur critique : Le label '" & label_to_find & "' est introuvable dans l'onglet '" & sheet_to_use.Name & "'.", vbCritical
        End
    End If
    Set letter_colonne = CelluleTrouvee
End Function


Function Plage_to_check(CelluleLabel As Range, sheet_to_use As Worksheet, table_name As String, start_row As Long) As Range
    Dim LettreColonne As String
    Dim DerniereLigne As Long
    
    ' On vérifie d'abord si la cellule titre existe
    If Not CelluleLabel Is Nothing Then
        LettreColonne = Split(CelluleLabel.Address, "$")(1)
        DerniereLigne = sheet_to_use.Cells(sheet_to_use.Rows.Count, LettreColonne).End(xlUp).Row
        
        ' On attribue l'objet final au nom de la fonction avec un SET
        Set Plage_to_check = sheet_to_use.Range(LettreColonne & start_row & ":" & LettreColonne & DerniereLigne)
    Else
        MsgBox "Le label '" & table_name & "' est introuvable sur la ligne 1.", vbCritical
        End
    End If
End Function


Function Get_sheet(table_name As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = Sheets(table_name)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "L'onglet '" & table_name & "' est introuvable.", vbCritical
        End
    End If
    Set Get_sheet = ws
End Function


Function get_path_directory(path_directory As String, path_directory_concat As String) As String
    Dim path_to_return As String
    Dim FSO As Object
    path_to_return = path_directory & "\" & path_directory_concat & "\"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(path_to_return) Then
        MsgBox "Erreur : Le dossier '" & path_to_return & "' est introuvable à l'emplacement :" & path_directory, vbCritical, "Dossier Manquant"
        End
    End If
    get_path_directory = path_to_return
End Function


Function get_data_dict() As Object
    Dim MonDico As Object
    Dim data_sheet As Worksheet
    Dim derniereColonne As Long
    
    Set data_sheet = Worksheets("DATA")
    Set MonDico = CreateObject("Scripting.Dictionary")
    derniereColonne = data_sheet.Cells(1, data_sheet.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To derniereColonne
        Cle = data_sheet.Cells(1, c).Value
        Valeur = data_sheet.Cells(2, c).Value
        If Cle <> "" Then
            MonDico(Cle) = Valeur
        End If
    Next c
    Set get_data_dict = MonDico
End Function

' =========================================================================
' 3. SUBPROCESS
' =========================================================================

Sub create_columns(lst_cols As Variant, sheet_new_cols As Worksheet)
    Dim derniereColonne As Long
    Dim colonneAInjecter As Long
    Dim NomColonne As Variant
    ' ajout nouvelle ligne pour noms colonnes
    sheet_new_cols.Rows(1).Insert Shift:=xlDown
    
    colonneAInjecter = 1
    
    For Each NomColonne In lst_cols
        sheet_new_cols.Cells(1, colonneAInjecter).Value = NomColonne
        colonneAInjecter = colonneAInjecter + 1
    Next NomColonne
End Sub


Sub inject_data_in_worbook(path_directory As String, sheet_to_create As String, ClasseurPrincipal As Workbook)
    Dim Dossier As Object
    Dim Fichier As Object
    Dim ClasseurSource As Workbook
    Dim FSO As Object
    Dim NouvelOnglet As Worksheet
    Dim FichierValideExiste As Boolean
    Dim wsAncienne As Worksheet
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Dossier = FSO.GetFolder(path_directory)
    
    FichierValideExiste = False
    
    ' Vérification de l'existence d'un fichier valide
    For Each Fichier In Dossier.Files
        If (InStr(Fichier.Name, ".xls") > 0) And (Left(Fichier.Name, 2) <> "~$") Then
            FichierValideExiste = True
            Exit For
        End If
    Next Fichier
    
    If Not FichierValideExiste Then
        MsgBox "Aucun fichier valide n'a été trouvé pour le traitement.", vbExclamation, "Dossier vide ou invalide"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each Fichier In Dossier.Files
        If (InStr(Fichier.Name, ".xls") > 0) And (Left(Fichier.Name, 2) <> "~$") Then
            ' Ouverture de la matrice en arrière-plan
            Set ClasseurSource = Workbooks.Open(Fichier.Path)
            
            ' 1. On supprime d'ABORD l'ancien onglet doublon s'il existait
            On Error Resume Next
            Set wsAncienne = ClasseurPrincipal.Worksheets(sheet_to_create)
            If Not wsAncienne Is Nothing Then wsAncienne.Delete
            On Error GoTo 0
            
            ' 2. CORRECTION : UN SEUL ".Add", placé directement au bon endroit
            Set NouvelOnglet = ClasseurPrincipal.Worksheets.Add(After:=ClasseurPrincipal.Sheets(ClasseurPrincipal.Sheets.Count))
            
            ' 3. On lui donne son nom
            NouvelOnglet.Name = sheet_to_create

            ' 4. CORRECTION COPIE : Utilisation de PasteSpecial (évite les plantages inter-classeurs)
            ClasseurSource.Sheets(1).UsedRange.Copy
            NouvelOnglet.Range("A1").PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False ' Vide le presse-papiers

            ' Fermeture de la matrice
            ClasseurSource.Close SaveChanges:=False
            
            ' Note : Si vous n'avez qu'un seul fichier à traiter dans le dossier, 
            ' décommentez la ligne ci-dessous pour éviter qu'un éventuel 2ème fichier n'écrase le 1er.
            ' Exit For
        End If
    Next Fichier
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub generate_matrice(Plage_libelle As Range, PlageRecherche As Range, Cellule_libelle_clean As Range, Cellule_key_table As Range, Cellule_IBAN As Range, Cellule_RIB As Range, Cellule_BIC As Range, Cellule_date_ech As Range, IBAN As String, RIB As String, BIC As String)
    Const val_not_in_correspondance As String = "##"
    Const char_to_split As String = " ech "
    Const intitule As String = "Intitulé"
    
    Dim texte_cell As String
    Dim Cellule As Range
    Dim TableauMorceaux() As String
    Dim PremierePartie As String
    Dim DateBrute As String
    Dim date_ech As Date
    Dim CleTrouvee As Range
    ' On mémorise la feuille pour que .Cells sache où écrire précisément
    Dim wsTarget As Worksheet
    Dim wsSearch As Worksheet
    Dim DateClean As String
    
    Set wsTarget = Plage_libelle.Worksheet
    Set wsSearch = PlageRecherche.Worksheet
    
    ' Optimisation de la matrice
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each Cellule In Plage_libelle
        texte_cell = Cellule.Value
        
        If texte_cell <> "" Then
            If InStr(texte_cell, char_to_split) > 1 Then
                TableauMorceaux = Split(texte_cell, char_to_split)
                PremierePartie = Trim(TableauMorceaux(0))
                DateBrute = Replace(Trim(TableauMorceaux(1)), ".", "/")
                DateClean = Trim(Split(DateBrute, " ")(0))
                
                If IsDate(DateClean) Then
                    date_ech = CDate(DateClean)
                Else
                    MsgBox "La date '" & date_ech & "' est mal formée pour le libellé '" & texte_cell & "', le format doit être 'jj.mm.yyyy'.", vbCritical, "Mauvais format date"
                    End
                End If
                
                ' On récupère le numéro de la ligne actuelle
                Dim r As Long
                r = Cellule.Row
                
                wsTarget.Cells(r, Cellule_libelle_clean.Column).Value = PremierePartie
                
                ' LE RAPPROCHEMENT
                Set CleTrouvee = PlageRecherche.Find(What:=PremierePartie, LookIn:=xlValues, LookAt:=xlWhole)
                
                If Not CleTrouvee Is Nothing Then
                    wsTarget.Cells(r, Cellule_key_table.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(intitule, wsSearch).Column).Value
                    wsTarget.Cells(r, Cellule_IBAN.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(IBAN, wsSearch).Column).Value
                    
                    wsTarget.Cells(r, Cellule_date_ech.Column).Value = date_ech
                
                    With wsTarget.Cells(r, Cellule_RIB.Column)
                        .NumberFormat = "@"
                        .Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(RIB, wsSearch).Column).Value
                    End With
                
                    wsTarget.Cells(r, Cellule_BIC.Column).Value = wsSearch.Cells(CleTrouvee.Row, letter_colonne(BIC, wsSearch).Column).Value
                Else
                    wsTarget.Cells(r, Cellule_key_table.Column).Value = val_not_in_correspondance
                End If
            End If
        End If
        
    Next Cellule
    
    ' Réactivation d'Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


Sub get_debit(Plage_debit As Range, PlageRecherche As Range, Cellule_libelle As Range, Cellule_IBAN As Range, Cellule_RIB As Range, Cellule_BIC As Range, Cellule_date_ech As Range, Cellule_key_table As Range, IBAN As String, RIB As String, BIC As String)
                
    ' si on a une valeur dans la colonne "débit" on déclenche la recherche du libellé correspondant, ex : BONDU AV2072131 SF71668
    ' on doit retrouver BONDU, on reprend l'onglet "table de correspondance" et la colonne "correspondance coala"
    ' on itère sur chaque valeur pour vérifier si la valeur de la cellule se retrouve dans notre libellé de Débit
    ' si on le retrouve, on injecte les valeurs bancaires associées
    Const char_to_split As String = " ech "
    
    Dim Cellule As Range
    Dim r As Long
    Dim correspondance_coala As Range
    Dim NomCorrespondance As String
    Dim texte_libelle As String
    Dim TableauMorceaux() As String
    Dim DateBrute As String
    Dim date_ech As Date
    Dim DateClean As String
    Dim search_IBAN As Range
    Dim search_RIB As Range
    Dim search_BIC As Range
    
    Set wsTarget = Plage_debit.Worksheet
    Set wsSearch = PlageRecherche.Worksheet
    
    Set search_IBAN = letter_colonne(IBAN, wsSearch)
    Set search_RIB = letter_colonne(RIB, wsSearch)
    Set search_BIC = letter_colonne(BIC, wsSearch)
    
    For Each Cellule In Plage_debit
        texte_cell = Cellule.Value
        
        r = Cellule.Row
        texte_libelle = Trim(CStr(wsTarget.Cells(r, Cellule_libelle.Column).Value))

        If texte_cell <> 0 Then
            'traitement : si débit différent de 0 on déclenche la recherche du libellé dans correspondance coala de notre table de correspondance
            For Each correspondance_coala In PlageRecherche
            
                NomCorrespondance = Trim(CStr(correspondance_coala.Value))
                If NomCorrespondance <> "" And NomCorrespondance <> "##" Then
                
                    If InStr(1, texte_libelle, NomCorrespondance) > 0 And InStr(texte_libelle, char_to_split) > 1 Then
                        TableauMorceaux = Split(texte_libelle, char_to_split)
                        DateBrute = Replace(Trim(TableauMorceaux(1)), ".", "/")
                        DateClean = Trim(Split(DateBrute, " ")(0))
                        
                        If IsDate(DateClean) Then
                            date_ech = CDate(DateClean)
                        Else
                            MsgBox "La date '" & date_ech & "' est mal formée pour le libellé '" & texte_libelle & "', le format doit être 'jj.mm.yyyy'.", vbCritical, "Mauvais format date"
                            End
                        End If
                    
                        wsTarget.Cells(r, Cellule_IBAN.Column).Value = wsSearch.Cells(correspondance_coala.Row, search_IBAN.Column).Value
                        
                        With wsTarget.Cells(r, Cellule_RIB.Column)
                            .NumberFormat = "@"
                            .Value = wsSearch.Cells(correspondance_coala.Row, search_RIB.Column).Value
                        End With
                        
                        wsTarget.Cells(r, Cellule_BIC.Column).Value = wsSearch.Cells(correspondance_coala.Row, search_BIC.Column).Value
                        
                        wsTarget.Cells(r, Cellule_key_table.Column).Value = NomCorrespondance
                        
                        wsTarget.Cells(r, Cellule_date_ech.Column).Value = date_ech
                        
                    End If
                End If
            Next correspondance_coala
        End If
        
    Next Cellule
    
End Sub


Sub generate_xml(CheminDossierOut As String, xml_name As String, export_achat As Worksheet, Cellule_date As Range, Cellule_libelle_clean As Range, start_row As Long, _
                credit As String, debit_lib As String, BIC_fournisseur As String, key_table As String, IBAN_fournisseur As String, RIB_fournisseur As String, _
                date_libelle As String, ByRef number_of_transactions As Long, ByRef total_payments As Double, ByRef paiement_dict As Object, DateDebutChoisie As Date, DateFinChoisie As Date, _
                Compte_export_achat As String, data_dict As Object)
    
    Const NS_PAIN001 As String = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.02"

    Dim ColBic As Long
    Dim ColKeyTable As Long
    Dim ColIBAN As Long
    Dim ColRIB As Long
    Dim ColDate As Long
    Dim Plage_date As Range
    Dim plage_cellule_rib As Range
    Dim texte_cell As String
    Dim r As Long
    Dim valeur_credit As Double
    Dim valeur_debit As Double
    Dim ColCreditNum As Long
    Dim ColDebitNum As Long
    Dim Id_transac As String
    Dim valeur_date As Date
    Dim DebitDico As Object
    Dim DebitKeyDico As Object
    Dim SousDico As Object
    Dim key_deb As String
    Dim debit As Variant
    Dim montantVirement As Double
    Dim TexteMessage As String
    Dim CheminFichier As String
    CheminFichier = CheminDossierOut & xml_name
    
    number_of_transactions = 0
    total_payments = 0
    
    Dim DocXml As Object
    Set DocXml = CreateObject("MSXML2.DOMDocument")
    
    ' Entête XML (Ligne obligatoire : <?xml version="1.0" encoding="UTF-8"?>)
    Set ProcInstr = DocXml.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    DocXml.appendChild ProcInstr
    
    ' Création de la balise Racine (Le conteneur principal)
    Set Racine = DocXml.createNode(1, "Document", NS_PAIN001)
    DocXml.appendChild Racine
    
    Dim pain As Object
    Set pain = DocXml.createNode(1, "pain.001.001.02", NS_PAIN001)
    Racine.appendChild pain
    
    Dim GrpHdr As Object
    Set GrpHdr = DocXml.createNode(1, "GrpHdr", NS_PAIN001)
    pain.appendChild GrpHdr
    
    
    Dim MsgId As Object
    Set MsgId = DocXml.createNode(1, "MsgId", NS_PAIN001)
    MsgId.Text = data_dict("MsgId_client") & Format(Now, "ddmmyyyy-hhmmss")
    GrpHdr.appendChild MsgId
    
    Dim CreDtTm As Object
    Set CreDtTm = DocXml.createNode(1, "CreDtTm", NS_PAIN001)
    CreDtTm.Text = Format(Now, "yyyy-mm-ddThh:mm:ss")
    GrpHdr.appendChild CreDtTm
    
    Dim NbOfTxs As Object
    Set NbOfTxs = DocXml.createNode(1, "NbOfTxs", NS_PAIN001)
    GrpHdr.appendChild NbOfTxs
    
    Dim CtrlSum As Object
    Set CtrlSum = DocXml.createNode(1, "CtrlSum", NS_PAIN001)
    GrpHdr.appendChild CtrlSum
    
    Dim Grpg As Object
    Set Grpg = DocXml.createNode(1, "Grpg", NS_PAIN001)
    Grpg.Text = data_dict("Grpg_client")
    GrpHdr.appendChild Grpg
    
    Dim InitgPty As Object
    Set InitgPty = DocXml.createNode(1, "InitgPty", NS_PAIN001)
    GrpHdr.appendChild InitgPty

    Dim Nm As Object
    Set Nm = DocXml.createNode(1, "Nm", NS_PAIN001)
    Nm.Text = data_dict("Nm_client")
    InitgPty.appendChild Nm
    
    Dim Id As Object
    Set Id = DocXml.createNode(1, "Id", NS_PAIN001)
    InitgPty.appendChild Id
    
    Dim OrgId As Object
    Set OrgId = DocXml.createNode(1, "OrgId", NS_PAIN001)
    Id.appendChild OrgId
    
    Dim PrtryId As Object
    Set PrtryId = DocXml.createNode(1, "PrtryId", NS_PAIN001)
    OrgId.appendChild PrtryId
    
    Dim IdSiret As Object
    Set IdSiret = DocXml.createNode(1, "Id", NS_PAIN001)
    IdSiret.Text = data_dict("Siret_client")
    PrtryId.appendChild IdSiret
    
    Dim PmtInf As Object
    Set PmtInf = DocXml.createNode(1, "PmtInf", NS_PAIN001)
    pain.appendChild PmtInf
    
    Dim PmtInfId As Object
    Set PmtInfId = DocXml.createNode(1, "PmtInfId", NS_PAIN001)
    Randomize
    PmtInfId.Text = data_dict("Nm_client") & Format(Now, "yymmddhhmmss") & Int(900 * Rnd + 100)
    PmtInf.appendChild PmtInfId
    
    Dim PmtMtd As Object
    Set PmtMtd = DocXml.createNode(1, "PmtMtd", NS_PAIN001)
    PmtMtd.Text = data_dict("PmtMtd")
    PmtInf.appendChild PmtMtd
    
    Dim PmtTpInf As Object
    Set PmtTpInf = DocXml.createNode(1, "PmtTpInf", NS_PAIN001)
    PmtInf.appendChild PmtTpInf
    
    Dim SvcLvl As Object
    Set SvcLvl = DocXml.createNode(1, "SvcLvl", NS_PAIN001)
    PmtTpInf.appendChild SvcLvl
    
    Dim Cd As Object
    Set Cd = DocXml.createNode(1, "Cd", NS_PAIN001)
    Cd.Text = data_dict("CdPmt")
    SvcLvl.appendChild Cd
    
    Dim ReqdExctnDt As Object
    Set ReqdExctnDt = DocXml.createNode(1, "ReqdExctnDt", NS_PAIN001)
    ReqdExctnDt.Text = Format(Now, "yyyy-mm-dd")
    PmtInf.appendChild ReqdExctnDt
    
    Dim Dbtr As Object
    Set Dbtr = DocXml.createNode(1, "Dbtr", NS_PAIN001)
    PmtInf.appendChild Dbtr
    
    Dim NmDbtr As Object
    Set NmDbtr = DocXml.createNode(1, "Nm", NS_PAIN001)
    NmDbtr.Text = data_dict("Nm_client")
    Dbtr.appendChild NmDbtr
    
    Dim PstlAdr As Object
    Set PstlAdr = DocXml.createNode(1, "PstlAdr", NS_PAIN001)
    Dbtr.appendChild PstlAdr
    
    Dim AdrLine As Object
    Set AdrLine = DocXml.createNode(1, "AdrLine", NS_PAIN001)
    AdrLine.Text = data_dict("AdrLine_client")
    PstlAdr.appendChild AdrLine
    
    Dim Ctry As Object
    Set Ctry = DocXml.createNode(1, "Ctry", NS_PAIN001)
    Ctry.Text = Left(data_dict("IBAN_client"), 2)
    PstlAdr.appendChild Ctry
    
    Dim DbtrAcct As Object
    Set DbtrAcct = DocXml.createNode(1, "DbtrAcct", NS_PAIN001)
    PmtInf.appendChild DbtrAcct
    
    Dim IdDbtr As Object
    Set IdDbtr = DocXml.createNode(1, "Id", NS_PAIN001)
    DbtrAcct.appendChild IdDbtr
    
    Dim IBAN As Object
    Set IBAN = DocXml.createNode(1, "IBAN", NS_PAIN001)
    IBAN.Text = data_dict("IBAN_client")
    IdDbtr.appendChild IBAN
    
    Dim DbtrAgt As Object
    Set DbtrAgt = DocXml.createNode(1, "DbtrAgt", NS_PAIN001)
    PmtInf.appendChild DbtrAgt
    
    Dim FinInstnId As Object
    Set FinInstnId = DocXml.createNode(1, "FinInstnId", NS_PAIN001)
    DbtrAgt.appendChild FinInstnId
    
    Dim BIC As Object
    Set BIC = DocXml.createNode(1, "BIC", NS_PAIN001)
    BIC.Text = data_dict("BIC_client")
    FinInstnId.appendChild BIC
    
    Dim ChrgBr As Object
    Set ChrgBr = DocXml.createNode(1, "ChrgBr", NS_PAIN001)
    ChrgBr.Text = data_dict("ChrgBr_client")
    PmtInf.appendChild ChrgBr
    
    Set Plage_date = Plage_to_check(Cellule_date, export_achat, export_achat.Name, start_row)
    
    ColBic = letter_colonne(BIC_fournisseur, export_achat).Column
    ColKeyTable = letter_colonne(key_table, export_achat).Column
    ColIBAN = letter_colonne(IBAN_fournisseur, export_achat).Column
    ColRIB = letter_colonne(RIB_fournisseur, export_achat).Column
    ColCreditNum = letter_colonne(credit, export_achat).Column
    ColDebitNum = letter_colonne(debit_lib, export_achat).Column
    
    ColDate = letter_colonne(date_libelle, export_achat).Column
    ColCompte = letter_colonne(Compte_export_achat, export_achat).Column
    Set DebitDico = CreateObject("Scripting.Dictionary")
    Set paiement_dict = CreateObject("Scripting.Dictionary")
	
	' on boucle une première fois pour obtenir l'intégralité des débits
	For Each plage_cellule_date In Plage_date
		r = plage_cellule_date.Row
		valeur_debit = CDbl(export_achat.Cells(r, ColDebitNum).Value)
		valeur_date = CDate(export_achat.Cells(r, ColDate).Value)
        valeur_compte = export_achat.Cells(r, ColCompte).Value
        valeur_RIB = export_achat.Cells(r, ColRIB).Value
		
		If valeur_RIB <> "" And valeur_debit <> 0 And valeur_date >= DateDebutChoisie And valeur_date <= DateFinChoisie And valeur_compte Like "F*" Then
            key_deb = export_achat.Cells(r, Cellule_libelle_clean.Column).Value
            
            Set DebitKeyDico = CreateObject("Scripting.Dictionary")
            DebitKeyDico("label") = export_achat.Cells(r, ColKeyTable).Value
            DebitKeyDico("montant") = valeur_debit
            Set DebitDico(key_deb) = DebitKeyDico
        End If
		
	Next plage_cellule_date
    
    ' boucle pour générer l'intégralité des virements à effectuer
    For Each plage_cellule_date In Plage_date
        r = plage_cellule_date.Row
        
        ' je caste en double
        valeur_credit = CDbl(export_achat.Cells(r, ColCreditNum).Value)

        ' On récupère la date de la ligne actuelle
        valeur_date = CDate(export_achat.Cells(r, ColDate).Value)
        valeur_compte = export_achat.Cells(r, ColCompte).Value
        valeur_RIB = export_achat.Cells(r, ColRIB).Value

        
        If valeur_RIB <> "" And valeur_credit <> 0 And valeur_date >= DateDebutChoisie And valeur_date <= DateFinChoisie And valeur_compte Like "F*" Then
        
            valeur_bic = export_achat.Cells(r, ColBic).Value
            valeur_KeyTable = export_achat.Cells(r, ColKeyTable).Value
            valeur_IBAN = export_achat.Cells(r, ColIBAN).Value
            
            Dim CdtTrfTxInf As Object
            Set CdtTrfTxInf = DocXml.createNode(1, "CdtTrfTxInf", NS_PAIN001)
            PmtInf.appendChild CdtTrfTxInf
        
            Dim PmtId As Object
            Set PmtId = DocXml.createNode(1, "PmtId", NS_PAIN001)
            CdtTrfTxInf.appendChild PmtId
            
            Dim InstrId As Object
            Set InstrId = DocXml.createNode(1, "InstrId", NS_PAIN001)
            PmtId.appendChild InstrId
            
            Dim EndToEndId As Object
            Set EndToEndId = DocXml.createNode(1, "EndToEndId", NS_PAIN001)
            PmtId.appendChild EndToEndId
            
            Dim Amt As Object
            Set Amt = DocXml.createNode(1, "Amt", NS_PAIN001)
            CdtTrfTxInf.appendChild Amt
            
            Dim InstdAmt As Object
            Set InstdAmt = DocXml.createNode(1, "InstdAmt", NS_PAIN001)
            InstdAmt.setAttribute "Ccy", "EUR"
            
            ' 1. On crée une variable de travail initialisée avec le montant du crédit d'origine
            Dim MontantRestantAPayer As Double
            MontantRestantAPayer = CDbl(valeur_credit)
            
            ' Sécurisation de la clé avec le numéro de ligne 'r' pour éviter les écrasements
            paiement_dict(valeur_KeyTable & "_" & r & "_" & Replace(CStr(MontantRestantAPayer), ",", ".")) = valeur_KeyTable & " : " & Replace(CStr(MontantRestantAPayer), ",", ".")
            
            Dim CleASupprimer As String
            CleASupprimer = ""
            
            For Each debit In DebitDico.Keys
                ' On compare avec notre variable évolutive MontantRestantAPayer
				
                If DebitDico(debit)("label") = valeur_KeyTable And DebitDico(debit)("montant") < MontantRestantAPayer Then
                    ' On déduit le montant de notre variable de travail
                    MontantRestantAPayer = MontantRestantAPayer - DebitDico(debit)("montant")
                    
                    ' On enregistre la ligne de déduction dans le dictionnaire de résumé
                    paiement_dict(valeur_KeyTable & "_deb_" & DebitDico(debit)("montant")) = valeur_KeyTable & " : -" & DebitDico(debit)("montant")
                    
                    CleASupprimer = debit
                    Exit For
                End If
            Next debit
            
            ' Suppression sécurisée après la boucle
            If CleASupprimer <> "" Then
                DebitDico.Remove CleASupprimer
            End If
            
            ' 2. C'est SEULEMENT ICI, à la fin, qu'on attribue le montant final (avec le point XML) au nœud XML
            InstdAmt.Text = Replace(CStr(MontantRestantAPayer), ",", ".")
            
            Amt.appendChild InstdAmt
            
            Dim CdtrAgt As Object
            Set CdtrAgt = DocXml.createNode(1, "CdtrAgt", NS_PAIN001)
            CdtTrfTxInf.appendChild CdtrAgt
            
            Dim FinInstnId_fournisseur As Object
            Set FinInstnId_fournisseur = DocXml.createNode(1, "FinInstnId", NS_PAIN001)
            CdtrAgt.appendChild FinInstnId_fournisseur
            
            Dim BIC_payment As Object
            Set BIC_payment = DocXml.createNode(1, "BIC", NS_PAIN001)
            BIC_payment.Text = valeur_bic
            FinInstnId_fournisseur.appendChild BIC_payment
            
            Dim Cdtr As Object
            Set Cdtr = DocXml.createNode(1, "Cdtr", NS_PAIN001)
            CdtTrfTxInf.appendChild Cdtr
            
            Dim Nm_fournisseur As Object
            Set Nm_fournisseur = DocXml.createNode(1, "Nm", NS_PAIN001)
            Nm_fournisseur.Text = valeur_KeyTable
            Cdtr.appendChild Nm_fournisseur
            
            Dim PstlAdr_fournisseur As Object
            Set PstlAdr_fournisseur = DocXml.createNode(1, "PstlAdr", NS_PAIN001)
            Cdtr.appendChild PstlAdr_fournisseur
            
            
            Dim Ctry_fournisseur As Object
            Set Ctry_fournisseur = DocXml.createNode(1, "Ctry", NS_PAIN001)
            Ctry_fournisseur.Text = Left(valeur_IBAN, 2)
            PstlAdr_fournisseur.appendChild Ctry_fournisseur
            
            Dim CdtrAcct_fournisseur As Object
            Set CdtrAcct_fournisseur = DocXml.createNode(1, "CdtrAcct", NS_PAIN001)
            CdtTrfTxInf.appendChild CdtrAcct_fournisseur
            
            Dim Id_fournisseur As Object
            Set Id_fournisseur = DocXml.createNode(1, "Id", NS_PAIN001)
            CdtrAcct_fournisseur.appendChild Id_fournisseur
            
            Dim iban_payment As Object
            Set iban_payment = DocXml.createNode(1, "IBAN", NS_PAIN001)
            iban_payment.Text = valeur_IBAN & valeur_RIB
            Id_fournisseur.appendChild iban_payment
            
            Dim RgltryRptg As Object
            Set RgltryRptg = DocXml.createNode(1, "RgltryRptg", NS_PAIN001)
            CdtTrfTxInf.appendChild RgltryRptg
            
            Dim RgltryDtls As Object
            Set RgltryDtls = DocXml.createNode(1, "RgltryDtls", NS_PAIN001)
            RgltryRptg.appendChild RgltryDtls
            
            Dim Cd_payment As Object
            Set Cd_payment = DocXml.createNode(1, "Cd", NS_PAIN001)
            Cd_payment.Text = data_dict("CdRgltry")
            RgltryDtls.appendChild Cd_payment
            
            Dim RmtInf As Object
            Set RmtInf = DocXml.createNode(1, "RmtInf", NS_PAIN001)
            CdtTrfTxInf.appendChild RmtInf
            
            Dim Ustrd As Object
            Set Ustrd = DocXml.createNode(1, "Ustrd", NS_PAIN001)
            RmtInf.appendChild Ustrd
            
            number_of_transactions = number_of_transactions + 1
            total_payments = total_payments + MontantRestantAPayer
            
            Id_transac = number_of_transactions & "_" & CStr(Fix(MontantRestantAPayer)) & "_" & Format(Now, "yyyymmdd_hhmmss")
            InstrId.Text = Id_transac
            Ustrd.Text = Id_transac
            EndToEndId.Text = Id_transac
            
        End If
    Next plage_cellule_date
    
    NbOfTxs.Text = CStr(number_of_transactions)
    CtrlSum.Text = Replace(Format(total_payments, "0.00"), ",", ".")
    
    If Not DebitDico Is Nothing And DebitDico.Count > 0 Then
        For Each debit In DebitDico.Keys
            Set SousDico = DebitDico(debit)
            TexteMessage = TexteMessage & "• Clé : " & debit & _
                           " | Label : " & SousDico("label") & _
                           " | Montant : " & Format(SousDico("montant"), "#,##0.00 €") & vbCrLf
        Next debit
        MsgBox "Le dictionnaire contient les débits suivants qui n'ont pas été déduits" & TexteMessage
    End If
    
    DocXml.Save CheminFichier
End Sub


Sub archive_file(directory_out_archive As String, CheminDossierIn As String, CheminDossierOut As String, xml_name As String)
    Const NomNouveauDossier As String = "xml_from"

    Dim CheminFichierXml As String
    Dim Dossier As Object
    Dim FSO As Object
    Dim Fichier As Object
    Dim CheminSource As String
    Dim CheminDestination As String
    Dim HorodatageDossier As String
    Dim NomNouveauDossierHorodate As String
    
    CheminFichierXml = CheminDossierOut & xml_name
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Dossier = FSO.GetFolder(CheminDossierIn)
    
    HorodatageDossier = Format(Now, "yyyymmdd_hhmmss")
    NomNouveauDossierHorodate = NomNouveauDossier & "-" & HorodatageDossier
    
    CheminDestination = directory_out_archive & NomNouveauDossierHorodate
    
    MkDir CheminDestination
    
    For Each Fichier In Dossier.Files
        CheminSource = Fichier.Path
        FSO.MoveFile Source:=CheminSource, Destination:=CheminDestination & "\" & Fichier.Name
    Next Fichier
    
    FSO.CopyFile Source:=CheminFichierXml, Destination:=CheminDestination & "\" & xml_name
    
End Sub

