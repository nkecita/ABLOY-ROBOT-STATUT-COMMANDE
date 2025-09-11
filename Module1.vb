Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml



Public Structure configRobot



    Dim databasefichet As String
    Dim databasefichettest As String
    Dim filesxml As String
    Dim filesxmlbackup As String

    Dim jourhisto As Integer
    Dim test As String
    Dim smtphost As String
    Dim smtpport As String
    Dim smtplogin As String
    Dim smtppassword As String
    Dim smtpemailto As String

End Structure

Module Module1

    Public configGene As configRobot


    Sub Main()
        Try


            ' premier commit
            ' Lecture du fichier de configuration XML
            Dim i As Integer
            Dim startupPath As String
            Dim dsConfig As New DataSet
            Dim ismail As Boolean
            Dim OkUpdate As Boolean

            OkUpdate = True

            startupPath = Assembly.GetExecutingAssembly().GetName().CodeBase

            startupPath = Path.GetDirectoryName(startupPath)
            Console.WriteLine("Lecture configuration : " + startupPath + "\" + "config.xml")
            dsConfig.ReadXml(startupPath + "\" + "config.xml")

            If dsConfig.Tables.Count = 1 Then
                With dsConfig.Tables(0).Rows(0)
                    Console.WriteLine("Lecture fichier de config")

                    Console.WriteLine("Lecture fichier de config : TEST")
                    configGene.test = .Item("TEST")
                    Console.WriteLine("Lecture fichier de config : DATABASE-FICHET")
                    configGene.databasefichet = .Item("DATABASE-FICHET")
                    Console.WriteLine("Lecture fichier de config : DATABASE-FICHET-TEST")
                    configGene.databasefichettest = .Item("DATABASE-FICHET-TEST")
                    Console.WriteLine("Lecture fichier de config : FILES-XML")
                    configGene.filesxml = .Item("FILES-XML")
                    Console.WriteLine("Lecture fichier de config : FILES-XML-BACKUP")
                    configGene.filesxmlbackup = .Item("FILES-XML-BACKUP")
                    Console.WriteLine("Lecture fichier de config : JOURHISTO")
                    configGene.jourhisto = .Item("JOURHISTO")
                    Console.WriteLine("Lecture fichier de config : SMTP")
                    configGene.smtphost = .Item("SMTP")
                    Console.WriteLine("Lecture fichier de config : SMTPLOGIN")
                    configGene.smtplogin = .Item("SMTPLOGIN")
                    Console.WriteLine("Lecture fichier de config : SMTPPASSWORD")
                    configGene.smtppassword = .Item("SMTPPASSWORD")
                    Console.WriteLine("Lecture fichier de config : EMAILTO")
                    configGene.smtpemailto = .Item("EMAIL-TO")


                End With
            Else
                Console.WriteLine("Problème fichier de config")
            End If


            'préparation des répertoire
            Console.WriteLine("Préparation des répertoires")
            If Directory.Exists(configGene.filesxmlbackup) = False Then
                Directory.CreateDirectory(configGene.filesxmlbackup)
            End If

            ' Si jour férie ou fermé alors le traitement ne fait rien
            Console.WriteLine("Vérification jour férié")
            If is_dayoff() Then
                Return
            End If

            ' Desturction des anciens répertoires pour éviter une surcharge disque inutile
            Console.WriteLine("Purge des répertoires")
            purge_repertoire()

            ' Traitement des fichiers XML
            Console.WriteLine("lecture des fichiers")
            '            Dim di As New IO.DirectoryInfo(configGene.filesxml)
            Console.WriteLine("lecture des fichiers phase 2")



            Dim MyFiles As String() = IO.Directory.GetFiles(configGene.filesxml, "*.xml")


            'Si fichiers trouvés
            If MyFiles IsNot Nothing Then
                Dim MySortedFiles As New List(Of String)
                MySortedFiles.Add(MyFiles(0))
                Console.WriteLine("Traitement fichier XML" + MyFiles(0))
                Console.WriteLine("Nombre de fichiers : " + MyFiles.Length.ToString)
                'Tri les fichiers par date du plus ancien au plus récent
                'Dans la collection MySortedFiles

                For z = 1 To MyFiles.Length - 1
                    Dim CurrentDate As Date = New IO.FileInfo(MyFiles(z)).CreationTime

                    For j As Integer = 0 To MySortedFiles.Count - 1
                        Dim NextDate As Date = New IO.FileInfo(MySortedFiles(j)).CreationTime
                        'Compare les dates du fichier précédent avec date du fichier en cours
                        If Date.Compare(CurrentDate, NextDate) < 0 Then
                            MySortedFiles.Insert(j, MyFiles(z))
                            Exit For
                        ElseIf j = MySortedFiles.Count - 1 Then
                            MySortedFiles.Add(MyFiles(z))
                            Exit For
                        End If

                    Next
                Next
            Else
                Console.WriteLine("Pas de fichiers XML dans le répertoire")
            End If






            Console.WriteLine("lecture des fichiers phase 3")

            Console.WriteLine("lecture des fichiers phase 4")

            Dim DsCommande As New DataSet
            Console.WriteLine("Lecture des commandes")
            For i = 0 To MyFiles.Count - 1
                Console.WriteLine("Lecture commande")
                DsCommande.ReadXml(MyFiles(i))
                Console.WriteLine("Lecture Fichier : " + MyFiles(i))
                If DsCommande.Tables.Count = 1 Then
                    For t = 0 To DsCommande.Tables(0).Rows.Count - 1
                        Console.WriteLine(DsCommande.Tables(0).Rows(t).Item("NUMCOM"))
                        Dim cdeEcon As String = DsCommande.Tables(0).Rows(t).Item("NUMCOM")
                        cdeEcon = Regex.Match(cdeEcon, "\d+").Value
                        ismail = False
                        With DsCommande.Tables(0).Rows(t)
                            OkUpdate = True
                            ismail = is_mail(CInt(cdeEcon.ToString).ToString.Trim)


                            Dim csql As String

                            Dim ccon As New OleDb.OleDbConnection


                            If configGene.test = "OUI" Then
                                ccon.ConnectionString = configGene.databasefichettest
                            Else
                                ccon.ConnectionString = configGene.databasefichet
                            End If



                            csql = "UPDATE commandes_portes set status=?,date_usine=?,nom_usine=?,montant_usine=?,date_tournee=?,heure_tournee=?,mail=? where num_commande=?"

                            Dim oSqlAdapter As New OleDb.OleDbDataAdapter
                            oSqlAdapter.UpdateCommand = New OleDb.OleDbCommand(csql, ccon)
                            ccon.Open()

                            oSqlAdapter.UpdateCommand.Parameters.Clear()
                            ' Statut de la commande
                            Dim dbParam_status As New OleDb.OleDbParameter
                            dbParam_status.ParameterName = "@status"
                            Select Case .Item("STATUT").ToString.ToUpper
                                Case "CONFIRME"
                                    dbParam_status.Value = "3"
                                Case "FABRIQUE"
                                    dbParam_status.Value = "4"
                                Case "EXPEDIE"
                                    dbParam_status.Value = "5"
                                Case "ANNULE"
                                    dbParam_status.Value = "-1"
                                Case "REFUSE"
                                    dbParam_status.Value = "2R"
                                Case Else
                                    OkUpdate = False
                                    'dbParam_status.Value = DBNull.Value
                            End Select

                            dbParam_status.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_status)

                            ' Date de livraison
                            Dim dbParam_DateLiv As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_DateLiv.ParameterName = "@date"
                            dbParam_DateLiv.Value = .Item("DATE").ToString.Trim
                            dbParam_DateLiv.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_DateLiv)

                            'nom_usine
                            Dim dbParam_nomusine As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_nomusine.ParameterName = "@nomusine"
                            dbParam_nomusine.Value = Remplace_car_speciaux(.Item("NOM").ToString.Trim)
                            dbParam_nomusine.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_nomusine)

                            'montant_usine
                            Dim dbParam_montantusine As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_montantusine.ParameterName = "@montantusine"
                            dbParam_montantusine.Value = .Item("MT").ToString.Replace(",", ".").ToString.Trim
                            dbParam_montantusine.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_montantusine)

                            'Date de la tournée
                            Dim dbParam_datetournee As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_datetournee.ParameterName = "@datetournee"
                            dbParam_datetournee.Value = .Item("DATE_TOURNEE").ToString.Trim
                            dbParam_datetournee.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_datetournee)


                            'Heure de la tournée
                            Dim dbParam_heuretournee As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_heuretournee.ParameterName = "@heuretournee"
                            dbParam_heuretournee.Value = Left(.Item("HEURE_TOURNEE").ToString.Trim, 5)
                            dbParam_heuretournee.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_heuretournee)


                            'Mail
                            Dim dbParam_mail As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_mail.ParameterName = "@mail"
                            If .Item("STATUT").ToString.ToUpper = "CONFIRME" Then
                                dbParam_mail.Value = True
                            Else
                                dbParam_mail.Value = False
                            End If
                            dbParam_mail.DbType = System.Data.DbType.Boolean
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_mail)

                            'Numéro de commande
                            Dim dbParam_commande As System.Data.IDataParameter = New OleDb.OleDbParameter
                            dbParam_commande.ParameterName = "@commande"
                            dbParam_commande.Value = CInt(cdeEcon.ToString).ToString.Trim
                            dbParam_commande.DbType = System.Data.DbType.[String]
                            oSqlAdapter.UpdateCommand.Parameters.Add(dbParam_commande)





                            Dim rowsaffected As Integer = 0

                            If OkUpdate Then
                                Console.WriteLine("Mise à jour du statut de la commande")
                                Try

                                    rowsaffected = oSqlAdapter.UpdateCommand.ExecuteNonQuery()

                                    'MessageBox.Show("Mise à jour réussie. Lignes affectées : " & rowsaffected.ToString())

                                Catch ex As SqlClient.SqlException
                                    ' Erreurs SQL spécifiques (ex: violation de clé, syntaxe SQL...)
                                    log("Erreur SQL : " & ex.Message)

                                Catch ex As Exception
                                    ' Toutes les autres erreurs .NET
                                    log("Erreur générale : " & ex.Message)
                                End Try
                                'Console.WriteLine("Mise à jour su statut de la commande : " + rowsaffected.ToString)
                            End If

                            If rowsaffected = 1 Then
                                Console.WriteLine("Envoi du mail")
                                ' Envoi du mail de confirmation

                                If .Item("STATUT").ToString.ToUpper = "CONFIRME" Then

                                    Dim dsCde As New DataSet

                                    Dim Chaine_Sql As String

                                    Chaine_Sql = "select num_client,transfert,num_usine,date_usine,nom_usine,montant_usine,societe,email,num_commande,ref_client,ref_valideur,souche from commandes_portes where num_commande = " & CInt(cdeEcon.ToString).ToString.Trim
                                    Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

                                    Dim adapter = New OleDb.OleDbDataAdapter(Chaine_Sql, cBase)
                                    adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                                    adapter.Fill(dsCde, "list_commande")







                                    Dim myorder As String = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("num_usine")), "", dsCde.Tables(0).Rows(0).Item("num_usine"))
                                    Dim montantusine As String = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("montant_usine")), "", dsCde.Tables(0).Rows(0).Item("montant_usine"))



                                    Dim mycustomermail As String
                                    Dim refclient As String

                                    Dim Dateexp As String = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("date_usine")), "", dsCde.Tables(0).Rows(0).Item("date_usine"))



                                    Dim cImagePays As String
                                    Dim cLanguepf As String
                                    Dim mycustomer As String
                                    Dim cExisteValideur As String

                                    cExisteValideur = get_valideur_info(dsCde.Tables(0).Rows(0).Item("num_client").ToString, "codeclient")

                                    If cExisteValideur <> "" Then
                                        mycustomer = cExisteValideur.ToString
                                        cLanguepf = get_customer_info(mycustomer, "langue")
                                        mycustomermail = get_customer_info(mycustomer, "email")
                                        refclient = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("ref_valideur")), "", dsCde.Tables(0).Rows(0).Item("ref_valideur"))
                                    Else
                                        cLanguepf = get_customer_info(dsCde.Tables(0).Rows(0).Item("num_client").ToString, "langue")
                                        mycustomermail = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("email")), "", dsCde.Tables(0).Rows(0).Item("email"))
                                        refclient = IIf(IsDBNull(dsCde.Tables(0).Rows(0).Item("ref_client")), "", dsCde.Tables(0).Rows(0).Item("ref_client"))
                                    End If

                                    Select Case cLanguepf
                                        Case "es-ES"
                                            cImagePays = "<img  src='http://econ.fichet-pointfort.fr/images/bann-mail-es.jpg' border='0' /><BR>"
                                        Case "it_IT"
                                            cImagePays = "<img src='http://econ.fichet-pointfort.fr/images/bann-mail-it.jpg' border='0' /><BR>"
                                        Case "pt-PT"
                                            cImagePays = "<img  src='http://econ.fichet-pointfort.fr/images/bann-mail-pt.jpg' border='0' /><BR>"
                                        Case "en-GB"
                                            cImagePays = "<img  src='http://econ.fichet-pointfort.fr/images/bann-mail-en.jpg' border='0' /><BR>"
                                        Case "nl-BE"
                                            cImagePays = "<img  src='http://econ.fichet-pointfort.fr/images/bann-mail-nl.jpg' border='0' /><BR>"
                                        Case Else
                                            cImagePays = "<img  src='http://econ.fichet-pointfort.fr/images/bann-mail-fr.jpg' border='0' /><BR>"
                                    End Select

                                    Dim mybody As String
                                    mybody = cImagePays
                                    mybody = mybody & traduction(cLanguepf, "MAIL", "CONFADM2")
                                    mybody = mybody & dsCde.Tables(0).Rows(0).Item("souche").ToString & dsCde.Tables(0).Rows(0).Item("num_commande").ToString.PadLeft(7, "0")
                                    mybody = mybody & "<br>"
                                    mybody = mybody & traduction(cLanguepf, "MAIL", "CONFADM4")
                                    mybody = mybody & refclient
                                    mybody = mybody & "<br>"
                                    mybody = mybody & traduction(cLanguepf, "MAIL", "CONFADM5")
                                    mybody = mybody & myorder
                                    mybody = mybody & "<br>"
                                    mybody = mybody & traduction(cLanguepf, "MAIL", "CONFADM6")
                                    mybody = mybody & montantusine
                                    mybody = mybody & "<br>"
                                    mybody = mybody & traduction(cLanguepf, "MAIL", "CONFADM7")
                                    mybody = mybody & Dateexp

                                    Dim Myfrom As String
                                    Myfrom = traduction(cLanguepf, "MAIL", "CONFADM3")

                                    Dim mysubject As String
                                    mysubject = traduction(cLanguepf, "MAIL", "CONFADM1") & myorder

                                    If mycustomermail.Trim <> "" And ismail Then
                                        ' EnvoiMail(Myfrom, mysubject, mybody, mycustomermail)
                                    End If





                                End If


                            End If
                            ccon.Close()
                        End With
                    Next
                End If

                Dim nomfic As String

                nomfic = MyFiles(i).Substring(MyFiles(i).LastIndexOf("\") + 1)

                If File.Exists(configGene.filesxmlbackup & "\" & nomfic) Then
                    Console.WriteLine("Destruction des fichiers XMl : " & nomfic)
                    File.Delete(configGene.filesxmlbackup & "\" & nomfic)
                End If
                Console.WriteLine("Déplacement des fichiers XMl :  " & nomfic)

                'fi.MoveTo(configGene.filesxmlbackup & "\" & nomfic)
                My.Computer.FileSystem.MoveFile(MyFiles(i), configGene.filesxmlbackup & "\" & nomfic)

            Next
            
        Catch ex As Exception
            log(ErrorToString)


        End Try





    End Sub
   

    Function is_dayoff()
        Dim dsDayOff As DataSet = New DataSet()
        Dim csql As String

        csql = "SELECT dt_ferie FROM jours_feries WHERE dt_ferie = convert(varchar,getdate(),112)"
        Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

        Dim adapter = New OleDb.OleDbDataAdapter(csql, cBase)
        adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
        adapter.Fill(dsDayOff, "list_dayoff")

        If dsDayOff.Tables(0).Rows.Count = 1 Then
            Return True
        Else
            Return False
        End If




    End Function
    Function purge_repertoire()


        Dim di As New IO.DirectoryInfo(configGene.filesxmlbackup)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
        Dim fi As IO.FileInfo
        Dim cDateRep, cFileSup As String
        For Each fi In aryFi


            cDateRep = Left(fi.Name, 11)
            cDateRep = Right(cDateRep, 6)
            cFileSup = cDateRep.Substring(0, 2) & "/"
            cFileSup = cFileSup & cDateRep.Substring(2, 2) & "/"
            cFileSup = cFileSup & cDateRep.Substring(4, 2)

            If DateDiff(DateInterval.Day, CDate(cFileSup), Today) > configGene.jourhisto Then
                fi.Delete()

            End If

        Next




    End Function

    Function Remplace_car_speciaux(ByVal wContenu As String) As String
        wContenu = wContenu.Replace("&lt;", "<")
        wContenu = wContenu.Replace("&gt;", ">")
        wContenu = wContenu.Replace("&apos;", "'")
        wContenu = wContenu.Replace("&quot;", """")


        wContenu = wContenu.Replace("&#115;", "é")
        wContenu = wContenu.Replace("&#138;", "Š")
        wContenu = wContenu.Replace("&#142;", "Ž")
        wContenu = wContenu.Replace("&#154;", "š")
        wContenu = wContenu.Replace("&#158;", "ž")
        wContenu = wContenu.Replace("&#32;", "é")
        wContenu = wContenu.Replace("&#99;", "é")
        wContenu = wContenu.Replace("&#186;", "º")
        wContenu = wContenu.Replace("&#170;", "ª")

        wContenu = wContenu.Replace("&#192;", "À")
        wContenu = wContenu.Replace("&#193;", "Á")
        wContenu = wContenu.Replace("&#194;", "Â")
        wContenu = wContenu.Replace("&#195;", "Ã")
        wContenu = wContenu.Replace("&#196;", "Ä")
        wContenu = wContenu.Replace("&#197;", "Å")
        wContenu = wContenu.Replace("&#198;", "Æ")
        wContenu = wContenu.Replace("&#199;", "Ç")
        wContenu = wContenu.Replace("&#200;", "È")
        wContenu = wContenu.Replace("&#201;", "É")
        wContenu = wContenu.Replace("&#202;", "Ê")
        wContenu = wContenu.Replace("&#203;", "Ë")
        wContenu = wContenu.Replace("&#204;", "Ì")
        wContenu = wContenu.Replace("&#205;", "Í")
        wContenu = wContenu.Replace("&#206;", "Î")
        wContenu = wContenu.Replace("&#207;", "Ï")
        wContenu = wContenu.Replace("&#208;", "Ð")
        wContenu = wContenu.Replace("&#209;", "Ñ")
        wContenu = wContenu.Replace("&#210;", "Ò")
        wContenu = wContenu.Replace("&#211;", "Ó")
        wContenu = wContenu.Replace("&#212;", "Ô")
        wContenu = wContenu.Replace("&#213;", "Õ")
        wContenu = wContenu.Replace("&#214;", "Ö")
        wContenu = wContenu.Replace("&#215;", "×")
        wContenu = wContenu.Replace("&#216;", "Ø")
        wContenu = wContenu.Replace("&#217;", "Ù")
        wContenu = wContenu.Replace("&#218;", "Ú")
        wContenu = wContenu.Replace("&#219;", "Û")
        wContenu = wContenu.Replace("&#220;", "Ü")
        wContenu = wContenu.Replace("&#221;", "Ý")
        wContenu = wContenu.Replace("&#222;", "Þ")
        wContenu = wContenu.Replace("&#223;", "ß")
        wContenu = wContenu.Replace("&#224;", "à")
        wContenu = wContenu.Replace("&#225;", "á")
        wContenu = wContenu.Replace("&#226;", "â")
        wContenu = wContenu.Replace("&#227;", "ã")
        wContenu = wContenu.Replace("&#228;", "ä")
        wContenu = wContenu.Replace("&#229;", "å")
        wContenu = wContenu.Replace("&#230;", "æ")
        wContenu = wContenu.Replace("&#231;", "ç")
        wContenu = wContenu.Replace("&#232;", "è")
        wContenu = wContenu.Replace("&#233;", "é")
        wContenu = wContenu.Replace("&#234;", "ê")
        wContenu = wContenu.Replace("&#235;", "ë")
        wContenu = wContenu.Replace("&#236;", "ì")
        wContenu = wContenu.Replace("&#237;", "í")
        wContenu = wContenu.Replace("&#238;", "î")
        wContenu = wContenu.Replace("&#239;", "ï")
        wContenu = wContenu.Replace("&#240;", "ð")
        wContenu = wContenu.Replace("&#241;", "ñ")
        wContenu = wContenu.Replace("&#242;", "ò")
        wContenu = wContenu.Replace("&#243;", "ó")
        wContenu = wContenu.Replace("&#244;", "ô")
        wContenu = wContenu.Replace("&#245;", "õ")
        wContenu = wContenu.Replace("&#246;", "ö")
        wContenu = wContenu.Replace("&#247;", "÷")
        wContenu = wContenu.Replace("&#248;", "ø")
        wContenu = wContenu.Replace("&#249;", "ù")
        wContenu = wContenu.Replace("&#250;", "ú")
        wContenu = wContenu.Replace("&#251;", "û")
        wContenu = wContenu.Replace("&#252;", "ü")
        wContenu = wContenu.Replace("&#253;", "ý")
        wContenu = wContenu.Replace("&#254;", "þ")
        wContenu = wContenu.Replace("&#255;", "ÿ")
        wContenu = wContenu.Replace("&#268;", "Č")


        wContenu = wContenu.Replace("&#270;", "Ď")
        wContenu = wContenu.Replace("&#327;", "Ň")
        wContenu = wContenu.Replace("&#352;", "Š")
        wContenu = wContenu.Replace("&#381;", "Ž")
        wContenu = wContenu.Replace("&#269;", "č")

        wContenu = wContenu.Replace("&#271;", "ď")
        wContenu = wContenu.Replace("&#314;", "Ĺ")
        wContenu = wContenu.Replace("&#317;", "Ľ")
        wContenu = wContenu.Replace("&#318;", "ľ")
        wContenu = wContenu.Replace("&#328;", "ň")
        wContenu = wContenu.Replace("&#340;", "Ŕ")
        wContenu = wContenu.Replace("&#341;", "ŕ")
        wContenu = wContenu.Replace("&#353;", "š")
        wContenu = wContenu.Replace("&#356;", "Ť")
        wContenu = wContenu.Replace("&#357;", "ť")
        wContenu = wContenu.Replace("&#382;", "ž")



        wContenu = wContenu.Replace("&#199;", "Ç")
        wContenu = wContenu.Replace("&#10;", vbCrLf)
        wContenu = wContenu.Replace("&amp;", "")
        Return wContenu
    End Function
    Function log(ByVal wErreur As String)
        '  Try
        'EnvoiMail("support@ciage.fr", "Message Alerte robot Statut commandes", wErreur, configGene.smtpemailto)

        ' Dim streamWrite As New IO.StreamWriter("log.txt", True)
        ' streamWrite.WriteLine("Date de Traitement : " & Date.Today.Day.ToString.PadLeft(2, "0") & "/" & Date.Today.Month.ToString.PadLeft(2, "0") & "/" & Date.Today.Year.ToString & " " & Now.Hour.ToString.PadLeft(2, "0") & ":" & Now.Minute.ToString.PadLeft(2, "0"))
        ' streamWrite.WriteLine("************************************************************************")
        '  streamWrite.WriteLine(wErreur)
        '  streamWrite.WriteLine("************************************************************************")

        '  streamWrite.Close()
        '  Catch ex As Exception
        ' End Try


    End Function
    Private Sub EnvoiMail(ByVal De As String, ByVal Sujet As String, ByVal Message As String, ByVal wDest As String)
        Dim m As New MailMessage
        Dim SMTP_SERV As New SmtpClient

        SMTP_SERV.Host = configGene.smtphost

        SMTP_SERV.Port = configGene.smtpport

        m.From = New MailAddress(De)
        m.Subject = Sujet


        Try
            m.Body = Message
            m.To.Add(wDest)
            m.IsBodyHtml = True
            ' m.To.Add("nkecita@ciage.fr")
            'm.To.Add("xxxxxxxx@xxxxxxxxx.com") 'pour envoyer a plusieurs destinataires


            ' authenticatin
            Dim basicAuthenticationInfo As New System.Net.NetworkCredential(configGene.smtplogin, configGene.smtppassword)


            'send the message
         


            SMTP_SERV.EnableSsl = True

            SMTP_SERV.UseDefaultCredentials = False
            SMTP_SERV.Credentials = basicAuthenticationInfo

            SMTP_SERV.Send(m)
        Catch ex As Exception
            log(ErrorToString)

        End Try

    End Sub

    Function is_mail(ByVal wNumCommande As String)
        Try
            Dim Chaine_Sql As String
            Dim dscde As New DataSet
            Dim i As Integer
            Chaine_Sql = "select mail from commandes_portes where num_commande = " & wNumCommande
            Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(Chaine_Sql, cBase)
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dscde, "list_commande")

            For i = 0 To dscde.Tables(0).Rows.Count = 1
                If dscde.Tables(0).Rows(0).Item("mail").ToString = "1" Then
                    Return False
                Else
                    Return True

                End If
            Next

        Catch
            log(ErrorToString)
            Return False
        End Try


    End Function
    Public Function get_valideur_info(ByVal wClient As String, ByVal wInfo As String) As String
        Try
            Dim dsValideur As New DataSet

            Dim cSql As String

            cSql = " SELECT signature.CodeValideur, clients_new.CodeClient, clients_new.RaisonSociale" & _
                                      " FROM signature INNER JOIN" & _
                                      " clients_new ON clients_new.CodeClient = signature.CodeValideur" & _
                                      " WHERE     (signature.CodeClient = '" & wClient & "')"

            Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(cSql, cBase)

            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsValideur, "list_valideur")

            If dsValideur.Tables(0).Rows.Count = 1 Then
                Return dsValideur.Tables(0).Rows(0).Item(wInfo).ToString
            Else
                Return ""
            End If
        Catch
            log(ErrorToString)
            Return False
        End Try
    End Function

    Public Function get_customer_info(ByVal wClient As String, ByVal wInfo As String) As String
        Try

            Dim dsdataSet As New DataSet

            Dim cSql As String

            cSql = "SELECT " & wInfo.Trim() & " From clients_new where codeclient='" & wClient.Trim() & "'"

            Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(cSql, cBase)

            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsdataSet, "list_dataset")

            If dsdataSet.Tables(0).Rows.Count = 1 Then
                Return dsdataSet.Tables(0).Rows(0).Item(0).ToString
            Else
                Return ""
            End If
        Catch
            log(ErrorToString)
            Return False
        End Try

        
    End Function

    Public Function traduction(ByVal wPays As String, ByVal wPage As String, ByVal wElement As String) As String
        Try
            Dim dsdataSet As New DataSet

            Dim cSql As String

            cSql = "SELECT * " & _
                "FROM pays_libelle" & " where codelangue=" & "'" & wPays.ToUpper() & "'" & " And " & _
                "nom_page=" & "'" & wPage & "'" & " And " & _
                "codelibelle=" & "'" & wElement & "'"


            Dim cBase As String = IIf(configGene.test = "OUI", configGene.databasefichettest, configGene.databasefichet)

            Dim adapter = New OleDb.OleDbDataAdapter(cSql, cBase)

            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsdataSet, "list_dataset")

            If dsdataSet.Tables(0).Rows.Count = 1 Then
                Return dsdataSet.Tables(0).Rows(0).Item(3).ToString
            Else
                Return ""
            End If

        Catch
            log(ErrorToString)
            Return False
        End Try
    End Function
End Module
