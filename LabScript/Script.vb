
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Net
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Drawing

Module Script

    Class Geral
        Public Shared ScriptTxt As String = ""
        Public Shared TratadoTxt As String = ""
        Public Shared Tmps As Dictionary(Of String, String) = Nothing
        Public Shared Lotes As Dictionary(Of String, System.Text.StringBuilder)
        Public Shared ArqLog As String = ""
        Public Shared Help As Boolean = False
        Public Shared Result As String = ""
        Public Shared SemConfirm As Boolean = False
        Public Shared EmailFrom As String = ""
        Public Shared EmailTo As String = ""
        Public Shared EmailSubject As String = ""
        Public Shared EmailServer As String = ""
        Public Shared Porta As String = "25"
        Public Shared Usuario As String = ""
        Public Shared Senha As String = ""
    End Class

    Public Function VersaoApl() As String
        Return "V" & Format(My.Application.Info.Version.Major, "00") & "." & Format(My.Application.Info.Version.Minor, "00") & "." & Format(My.Application.Info.Version.MajorRevision, "00") & "." & Format(My.Application.Info.Version.MinorRevision, "00")
    End Function

    Sub Main()
        Try

            Geral.Result = ""

            Dim Args As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs()

            For z As Integer = 0 To Args.Count - 1 Step 2
                Dim Coma As String = SemAspas(Args(z)).ToLower
                Select Case Coma
                    Case "-semconfirm"
                        Geral.SemConfirm = True
                        z -= 1
                    Case "-arq"
                        Geral.ScriptTxt = CarregaArqTxt(SemAspas(Args(z + 1)))
                    Case "-de"
                        Geral.EmailFrom = SemAspas(Args(z + 1))
                    Case "-para"
                        Geral.EmailTo = SemAspas(Args(z + 1))
                    Case "-assunto"
                        Geral.EmailSubject = SemAspas(Args(z + 1))
                    Case "-servidor"
                        Geral.EmailServer = SemAspas(Args(z + 1))
                    Case "-porta"
                        Geral.Porta = SemAspas(Args(z + 1))
                    Case "-usuario"
                        Geral.Usuario = SemAspas(Args(z + 1))
                    Case "-senha"
                        Geral.Senha = SemAspas(Args(z + 1))
                    Case "-help"
                        Geral.Help = True
                        z -= 1
                    Case "-log"
                        Geral.ArqLog = SemAspas(Args(z + 1))
                End Select
            Next

            Dim Msg As String = vbCrLf & vbCrLf & "LabScript - " & VERSAOAPL & " - Laboratorio de script - Intercraft Solutions - 2012" & vbCrLf & vbCrLf
            If Geral.Help Or Args.Count = 0 Then

                Msg &= "     Modo de usar (exemplo)......................................." & vbCrLf
                Msg &= "     mostrar help:               -help" & vbCrLf
                Msg &= "     arquivo:                    -arq ""c:\script.lab""" & vbCrLf
                Msg &= "     enviar email de:            -de ""de@icraft.com.br""" & vbCrLf
                Msg &= "     enviar email para:          -para ""para@icraft.com.br""" & vbCrLf
                Msg &= "     assunto da mensagem:        -assunto ""LabScript Teste""" & vbCrLf
                Msg &= "     servidor:                   -servidor ""smtpi.icraft.com.br""" & vbCrLf
                Msg &= "     gravar log em (arquivo):    -log ""c:\labscript.log""" & vbCrLf
                Msg &= "     porta do servidor:          -porta" & vbCrLf
                Msg &= "     usuário de autenticacao:    -usuario" & vbCrLf
                Msg &= "     senha de autenticacao:      -senha" & vbCrLf
                Msg &= "     sem confirmacao:            -semconfirm" & vbCrLf & vbCrLf
                Msg &= "     Nos scripts podera utilizar:" & vbCrLf
                Msg &= "     ==yyyy        ano" & vbCrLf
                Msg &= "     ==mm          mes" & vbCrLf
                Msg &= "     ==dd          dia" & vbCrLf
                Msg &= "     ==hh          hora" & vbCrLf
                Msg &= "     ==mi          minutos" & vbCrLf
                Msg &= "     ==ss          segundos" & vbCrLf
                Msg &= vbCrLf
                Msg &= "     ==ultimodirn   busca ultimo dir naquele diretorio" & vbCrLf
                Msg &= vbCrLf
                Msg &= "     Ex.: 'C:\oracle\product\10.2.0\flash_recovery_area\SBDB\AUTOBACKUP\==ultimodirn\==ultimoarqc'" & vbCrLf
                Msg &= vbCrLf
                Msg &= "     ==ultimoarqc   busca ultimo arquivo naquele diretorio" & vbCrLf
                Msg &= vbCrLf
                Msg &= "     Ex.: COPY \\Oraclerjbkp\d$\BACKUP\==ultimoarqc D:\" & vbCrLf
                Msg &= vbCrLf
                Msg &= "     ==dskspace ou ==dskspaceC ou ==dskspaceD < espaço em disco" & vbCrLf
                Msg &= "     ==dskfree ou ==dskfreeC ou ==dskfreeD < espaço livre" & vbCrLf
                Msg &= vbCrLf
                Msg &= "Exemplo para mostrar tamanho e espaco livre:" & vbCrLf
                Msg &= "echo ==yyyy-==mm-==dd ==hh:==mi:==ss > c:\teste.log" & vbCrLf
                Msg &= "echo ==dskspace >> c:\teste.log" & vbCrLf
                Msg &= "echo ==dskfree >> c:\teste.log" & vbCrLf
                Msg &= vbCrLf
                RegLogLine(Msg)
                Exit Sub
            End If

            RegLogLine(Msg)

            Trata()

            RegLogLine(vbCrLf & "Arquivo tratado......................" & vbCrLf)
            RegLogLine(Geral.TratadoTxt)

            Continuar()

            AtribuiNomeArq()

            Dim ListaSeq() As String = (From x As String In Geral.Tmps.Keys Where Not System.Text.RegularExpressions.Regex.Match(x, "[^0-9]").Success Order By Val(x) Select x).ToArray
            For Each Item As String In ListaSeq
                Dim Iniciado As Date = Now
                Trata(Geral.Tmps)
                GravaArq()

                RegLogLine(vbCrLf & vbCrLf & "Executar lote [" & Item & "]")
                Continuar()

                Dim Erros As String = ""
                DosShell(Geral.Tmps(Item), "", "", "", Erros)

                Dim Termino As Date = Now
                RegLogLine(vbCrLf & "Inicio.: " & Format(Iniciado, "yyyy-MM-dd HH:mm:ss"))
                RegLogLine("Termino: " & Format(Termino, "yyyy-MM-dd HH:mm:ss"))
                RegLogLine("Duracao: " & DateDiff(DateInterval.Minute, Iniciado, Termino) & " mins" & vbCrLf)
            Next

            RegLogLine("Processo concluido.")

            If Geral.EmailFrom <> "" And Geral.EmailTo <> "" And Geral.EmailSubject <> "" Then
                Dim REt As String = EnviaEmail(Nothing, Nothing, Geral.EmailFrom, Geral.EmailTo, Geral.EmailSubject, Entifica(Geral.Result, TipoEntifica.Tudo).Replace(vbCrLf, "<br/>"), "", MailPriority.Normal, Geral.EmailServer, Geral.Porta, , , Geral.Usuario, Geral.Senha)
                RegLogLine("Mensagem enviada: " & NZV(REt, "OK"))
            End If

            If Geral.ArqLog <> "" Then
                Dim Arq As New System.IO.StreamWriter(Geral.ArqLog)
                Arq.WriteLine(Geral.Result)
                Arq.Close()
                RegLogLine("Log gravado em " & Geral.ArqLog)
            End If

            Continuar("Término do programa. Aperte qualquer tecla para continuar.")
        Catch ex As Exception
            RegLogLine("Erro " & ex.Message & " em LABSCRIPT.")
        End Try

        For Each Chave As String In Geral.Tmps.Keys
            Try
                Kill(Geral.Tmps(Chave))
            Catch
            End Try
        Next
    End Sub

    Dim TrataArqs As Boolean = False
    Sub GravaArq()
        If Not TrataArqs Then
            RegLogLine(vbCrLf & "Arquivos gravados (temps).............." & vbCrLf)
        End If
        Dim Pos As Integer = 0
        For Each Chave As String In Geral.Tmps.Keys
            Pos += 1
            If Not TrataArqs Then
                RegLogLine(Pos & "-" & Chave)
            End If
            Dim Tmp As String = Geral.Tmps(Chave)
            Dim F As New System.IO.StreamWriter(Tmp)
            F.Write(Geral.Lotes(Chave))
            F.Close()
            If Not TrataArqs Then
                RegLogLine("---> " & Tmp & vbCrLf)
            End If
        Next
        TrataArqs = True
    End Sub

    Sub Trata(Optional ByVal ArqLotes As Dictionary(Of String, String) = Nothing)
        Try
            BuscaLotes(Geral.ScriptTxt)
            AcertaLotes(ArqLotes)
            Dim Ret As String = ""

            For Each Chave As String In (From x As String In Geral.Lotes.Keys Order By x)
                Dim NomeArqLote As String = ""
                If Not IsNothing(ArqLotes) Then
                    If ArqLotes.ContainsKey(Chave) Then
                        NomeArqLote = ArqLotes(Chave)
                    End If
                End If

                Ret &= IIf(Ret <> "", vbCrLf & vbCrLf & "--------------------------------------------------------------" & vbCrLf, "") & "[" & Chave & "]" & IIf(NomeArqLote = "", "", " ====> " & NomeArqLote)
                Ret &= vbCrLf

                Ret &= Geral.Lotes(Chave).ToString
            Next

            Geral.TratadoTxt = Ret
        Catch Ex As Exception
            Throw New Exception(Ex.Message & " em tratar.")
        End Try
    End Sub

    Function SemAspas(ByVal Texto As String) As String
        Return Texto.Trim("""", Chr(147), Chr(148))
    End Function

    Public Function CarregaArqTxt(ByVal Arquivo As String) As String
        Dim BR As System.IO.BinaryReader = Nothing
        Dim FS As New System.IO.FileStream(Arquivo, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        BR = New System.IO.BinaryReader(FS)
        Return BR.ReadChars(Convert.ToInt32(BR.BaseStream.Length))
    End Function

    Sub BuscaLotes(ByVal Texto As String)
        Geral.Lotes = New Dictionary(Of String, System.Text.StringBuilder)

        Dim LoteAtual As String = ""
        For Each Linha As String In Split(Texto, vbCrLf)
            Dim Ctl As Boolean = False
            Linha = Trim(Linha)

            Dim M As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Linha, "^\[(.*?)\]$")
            If M.Success Then
                LoteAtual = M.Groups(1).Value
                Ctl = True
            Else
                If LoteAtual <> "" Then
                    M = System.Text.RegularExpressions.Regex.Match(Linha, "^\[\\" & LoteAtual & "\]$")
                    If M.Success Then
                        LoteAtual = ""
                        Ctl = True
                    End If
                End If
            End If

            ' quando não é controle
            If Not Ctl Then
                If Not Geral.Lotes.ContainsKey(LoteAtual) Then
                    Geral.Lotes.Add(LoteAtual, New System.Text.StringBuilder(Linha & vbCrLf))
                Else
                    Geral.Lotes(LoteAtual).AppendLine(Linha)
                End If
            End If
        Next
    End Sub

    Sub AcertaLotes(Optional ByVal ArqLotes As Dictionary(Of String, String) = Nothing)
        For Each Chave As String In (From x As String In Geral.Lotes.Keys Order By x)
            Dim ArqLote As String = TrataLote(Geral.Lotes(Chave).ToString)
            If Not IsNothing(ArqLotes) Then
                For Each M As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(ArqLote, "==\[(.*?)\]")
                    If ArqLotes.ContainsKey(M.Groups(1).Value) Then
                        Troca(ArqLote, M.Groups(0).Value, ArqLotes(M.Groups(1).Value))
                    End If
                Next
            End If
            Geral.Lotes(Chave) = New System.Text.StringBuilder(ArqLote)
        Next
    End Sub

    Function DskSpace(ByVal Disco As String) As String
        Dim DI As New Scripting.FileSystemObject
        Dim D As Scripting.Drive = DI.GetDrive(Disco & ":")
        Dim tot As Long = D.TotalSize

        Return tot & "byte" & IIf(tot <> 1, "s", "")
    End Function

    Function DskSpaceTodos() As String
        Dim Ret As String = ""
        For Z As Integer = Asc("A") To Asc("Z")
            Try
                Ret &= IIf(Ret <> "", " ", "") & "Dskspace" & Chr(Z) & "=" & DskSpace(Chr(Z))
            Catch
            End Try
        Next
        Return Ret
    End Function

    Function DskFree(ByVal Disco As String) As String
        Dim DI As New Scripting.FileSystemObject
        Dim D As Scripting.Drive = DI.GetDrive(Disco & ":")
        Dim tot As Long = D.AvailableSpace
        Return tot & "byte" & IIf(tot <> 1, "s", "")
    End Function

    Function DskFreeTodos() As String
        Dim Ret As String = ""
        For Z As Integer = Asc("A") To Asc("Z")
            Try
                Ret &= IIf(Ret <> "", " ", "") & "DskFree" & Chr(Z) & "=" & DskFree(Chr(Z))
            Catch
            End Try
        Next
        Return Ret
    End Function


    Function TrataLote(ByVal Texto As String) As String
        Dim Ret As String = Texto
        Dim Ret2 As String = ""
        Troca(Ret, "==yyyy", Format(Now, "yyyy"))
        Troca(Ret, "==mm", Format(Now, "MM"))
        Troca(Ret, "==dd", Format(Now, "dd"))
        Troca(Ret, "==hh", Format(Now, "HH"))
        Troca(Ret, "==mi", Format(Now, "mm"))
        Troca(Ret, "==ss", Format(Now, "ss"))

        For Each M As System.Text.RegularExpressions.Match In Regex.Matches(Ret, "==dskspace([a-z])", RegexOptions.IgnoreCase)
            Try
                Troca(Ret, M.Groups(0).Value, DskSpace(M.Groups(1).Value))
            Catch ex As Exception
                Troca(Ret, M.Groups(0).Value, "[ERRO] " & ex.Message)
            End Try
        Next
        If InStr(Ret, "==dskspace") <> 0 Then
            Troca(Ret, "==dskspace", DskSpaceTodos)
        End If

        For Each M As System.Text.RegularExpressions.Match In Regex.Matches(Ret, "==dskfree([a-z])", RegexOptions.IgnoreCase)
            Try
                Troca(Ret, M.Groups(0).Value, DskFree(M.Groups(1).Value))
            Catch ex As Exception
                Troca(Ret, M.Groups(0).Value, "[ERRO] " & ex.Message)
            End Try
        Next
        If InStr(Ret, "==dskfree") <> 0 Then
            Troca(Ret, "==dskfree", DskFreeTodos)
        End If

        For Each LINHA As String In Split(Ret, vbCrLf)
            LINHA = TrocaUltimoDir(LINHA, "c")
            LINHA = TrocaUltimoDir(LINHA, "n")
            LINHA = TrocaUltimoArq(LINHA, "c")
            LINHA = TrocaUltimoArq(LINHA, "n")
            Ret2 &= IIf(Ret2 <> "", vbCrLf, "") & LINHA
        Next
        Return Ret2
    End Function

    Function TrocaUltimoDir(ByVal Linha As String, Optional ByVal Ord As String = "")
        Dim Coma As String = "==ultimodir" & Ord
        Dim Pos As Integer = InStr(Linha, Coma)
        While Pos > 0
            Dim PosI1 As Integer = InStrRev(Linha, "'", Pos - 1)
            Dim POSI2 As Integer = InStrRev(Linha, """", Pos - 1)
            Dim PosI As Integer = Math.Max(PosI1, POSI2)
            If PosI = 0 Then
                Pos = InStr(Pos + Len(Coma), Linha, Coma)
            Else
                Dim DD As String = Mid(Linha, PosI + 1, Pos - PosI - 1)
                Dim DG() As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(DD).GetDirectories
                If DG.Count = 0 Then
                    Pos = InStr(Pos + Len(Coma), Linha, Coma)
                Else
                    Dim G As System.IO.DirectoryInfo = Nothing
                    If Ord = "c" Then
                        G = (From X As System.IO.DirectoryInfo In DG Order By X.CreationTime, X.Name Select X).Last
                    Else
                        G = (From X As System.IO.DirectoryInfo In DG Order By X.Name Select X).Last
                    End If
                    Linha = Microsoft.VisualBasic.Left(Linha, Pos - 1) & G.Name & Mid(Linha, Pos + Len(Coma))
                    Pos = InStr(Pos + Len(G.Name), Linha, Coma)
                End If
            End If
        End While
        Return Linha
    End Function

    Function TrocaUltimoArq(ByVal Linha As String, Optional ByVal Ord As String = "")
        Dim Coma As String = "==ultimoarq" & Ord
        Dim Pos As Integer = InStr(Linha, Coma)
        While Pos > 0
            Dim PosI1 As Integer = InStrRev(Linha, "'", Pos - 1)
            Dim POSI2 As Integer = InStrRev(Linha, """", Pos - 1)
            Dim PosI As Integer = Math.Max(PosI1, POSI2)
            If PosI = 0 Then
                Pos = InStr(Pos + Len(Coma), Linha, Coma)
            Else
                Dim DD As String = Mid(Linha, PosI + 1, Pos - PosI - 1)
                Dim DG() As System.IO.FileInfo = New System.IO.DirectoryInfo(DD).GetFiles
                If DG.Count = 0 Then
                    Pos = InStr(Pos + Len(Coma), Linha, Coma)
                Else
                    Dim G As System.IO.FileInfo = Nothing
                    If Ord = "c" Then
                        G = (From X As System.IO.FileInfo In DG Order By X.CreationTime, X.Name Select X).Last
                    Else
                        G = (From X As System.IO.FileInfo In DG Order By X.Name Select X).Last
                    End If
                    Linha = Microsoft.VisualBasic.Left(Linha, Pos - 1) & G.Name & Mid(Linha, Pos + Len(Coma))
                    Pos = InStr(Pos + Len(G.Name), Linha, Coma)
                End If
            End If
        End While
        Return Linha
    End Function

    Sub Troca(ByRef Fonte As String, ByVal De As String, ByVal Para As String)
        Dim Pos As Integer = InStr(Fonte, De)
        Do While Pos <> 0
            Fonte = Microsoft.VisualBasic.Left(Fonte, Pos - 1) & Para & Mid(Fonte, Pos + Len(De))
            Pos = InStr(Pos + Len(Para), Fonte, De)
        Loop
    End Sub

    Sub AtribuiNomeArq()
        Geral.Tmps = New Dictionary(Of String, String)
        For Each Chave As String In Geral.Lotes.Keys
            Dim Tmp As String = System.IO.Path.GetTempFileName
            If Not System.Text.RegularExpressions.Regex.Match(Chave, "[^0-9]").Success Then
                Dim SEQ As Integer = 0
                Do While True
                    Dim NomeBAt As String = Tmp & IIf(SEQ = 0, "", SEQ) & ".BAT"
                    If Not System.IO.File.Exists(NomeBAt) Then
                        System.IO.File.Move(Tmp, NomeBAt)
                        Tmp = NomeBAt
                        Exit Do
                    End If
                    SEQ += 1
                Loop
            End If
            Geral.Tmps.Add(Chave, Tmp)
        Next
    End Sub

    Sub Continuar(Optional ByVal Texto As String = "Deseja continuar (s/n)?")
        If Not Geral.SemConfirm Then
            RegLogLine(vbCrLf)
            RegLog(Texto)
            If System.Console.ReadKey().Key <> ConsoleKey.S Then
                End
            End If
            RegLogLine(vbCrLf)
        End If

    End Sub

    Public Function DosShell(ByVal Comando As String, ByVal Argumento As String, ByVal Diretorio As String, ByVal Entrada As String, ByRef Erros As String, Optional ByVal Usuario As String = "", Optional ByVal Senha As String = "", Optional ByVal Dominio As String = "", Optional ByVal EsperaSegs As Integer = 30) As String
        Dim Result As String = ""
        Try
            Dim Proc As System.Diagnostics.Process = Nothing
            Dim StdIn As System.IO.StreamWriter = Nothing
            Dim StdOut As System.IO.StreamReader = Nothing
            Dim StdErr As System.IO.StreamReader = Nothing

            If Diretorio = "" Then
                Diretorio = System.IO.Path.GetTempPath()
            End If

            Dim Psi As New System.Diagnostics.ProcessStartInfo(Comando, Argumento)
            Psi.CreateNoWindow = True
            Psi.ErrorDialog = False
            Psi.UseShellExecute = False
            Psi.RedirectStandardError = True
            Psi.RedirectStandardInput = True
            Psi.RedirectStandardOutput = True
            Psi.WorkingDirectory = Diretorio

            If Usuario <> "" Then
                Psi.UserName = Usuario
            End If
            If Senha <> "" Then
                Psi.Password = New System.Security.SecureString
                For z As Integer = 1 To Len(Senha)
                    Psi.Password.AppendChar(Mid(Senha, z, 1))
                Next
            End If
            If Dominio <> "" Then
                Psi.Domain = Dominio
            End If

            Proc = System.Diagnostics.Process.Start(Psi)

            StdIn = Proc.StandardInput
            StdOut = Proc.StandardOutput
            StdErr = Proc.StandardError

            If Entrada <> "" Then
                StdIn.WriteLine(Entrada)
            End If

            Do While Not Proc.HasExited
                Dim buf As String
                buf = StdOut.ReadLine()
                RegLogLine(buf)
                Result &= buf
            Loop
            RegLog(vbCrLf & "ExitCode:" & Proc.ExitCode)
            Erros = StdErr.ReadToEnd()
            If Erros <> "" Then
                RegLog(vbCrLf & vbCrLf & "Erros:" & vbCrLf)
                RegLog(Erros)
            End If
        Catch ex As Exception
            Result = "Erro ao executar comando: " & ex.Message
        End Try
        Return Result
    End Function

    Sub RegLog(ByVal Texto As String)
        System.Console.Write(Texto)
        Geral.Result &= Texto
    End Sub

    Sub RegLogLine(ByVal Texto As String)
        System.Console.WriteLine(Texto)
        Geral.Result &= Texto & vbCrLf
    End Sub

    Public Function EnviaEmail(ByRef Mail As MailMessage, ByVal Enviar As System.Net.Mail.SmtpClient, ByVal De As String, ByVal Para As Object, ByVal Assunto As String, ByVal Corpo As String, ByVal ReplyTo As String, Optional ByVal Prioridade As System.Net.Mail.MailPriority = Nothing, Optional ByVal SmtpHost As String = Nothing, Optional ByVal SmtpPort As Integer = 25, Optional ByVal CC As Object = Nothing, Optional ByVal BCC As Object = Nothing, Optional ByVal SMTPUsuario As String = Nothing, Optional ByVal SMTPSenha As String = Nothing, Optional ByVal IncorporaImagens As Boolean = False, Optional ByRef CIDS As ArrayList = Nothing, Optional ByRef TMPS As ArrayList = Nothing, Optional ByVal UrlsLocais As ArrayList = Nothing, Optional ByVal Attachs As ArrayList = Nothing) As String
        Try


            ' cada param só é definido caso esteja mencionado
            If IsNothing(Mail) Then
                Mail = New MailMessage
            End If

            If Not IsNothing(De) Then
                Dim DeLista As ArrayList = TermosStrToLista(De)
                Mail.From = New MailAddress(EmailStr(DeLista(0)))
            End If

            If Not IsNothing(ReplyTo) Then
                Dim ReplyToLista As ArrayList = TermosStrToLista(ReplyTo)
                If ReplyToLista.Count > 0 Then
                    Mail.ReplyTo = New MailAddress(EmailStr(ReplyToLista(0)))
                End If
            End If

            If Not IsNothing(Para) Or Not IsNothing(CC) Or Not IsNothing(BCC) Then
                Mail.Bcc.Clear()
                Mail.CC.Clear()
                Mail.To.Clear()
            End If

            If Not IsNothing(Para) Then
                Dim ParaLista As ArrayList = TermosStrToLista(Para)
                For Each ParaItem As String In ParaLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Mail.To.Add(New MailAddress(EmailStr(ParaItem)))
                    End If
                Next
            End If

            If Not IsNothing(CC) Then
                Dim CCLista As ArrayList = TermosStrToLista(CC)
                For Each ParaItem As String In CCLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Mail.CC.Add(New MailAddress(EmailStr(ParaItem)))
                    End If
                Next
            End If

            If Not IsNothing(BCC) Then
                Dim BCCLista As ArrayList = TermosStrToLista(BCC)
                For Each ParaItem As String In BCCLista
                    If ParaItem.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        Dim M As New Email(ParaItem.Substring(4))
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    Else
                        Dim M As New Email(ParaItem)
                        Mail.Bcc.Add(New MailAddress("<" & M.SoEndereco & ">"))
                    End If
                Next
            End If

            If Not IsNothing(Prioridade) Then
                Mail.Priority = Prioridade
            End If

            If Not IsNothing(Assunto) Then
                Mail.Subject = Assunto
            End If

            If Not IsNothing(Corpo) Then
                Mail.AlternateViews.Clear()


                If Not IncorporaImagens Then
                    Mail.IsBodyHtml = True
                    Mail.SubjectEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                    Mail.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8")
                    Mail.Body = Corpo
                Else

                    ' inicia variávies de retorno caso não estejam definidas
                    If IsNothing(TMPS) Then
                        TMPS = New ArrayList
                    End If
                    If IsNothing(CIDS) Then
                        CIDS = New ArrayList
                    End If

                    ' visão alternativa
                    Dim alt As AlternateView = AlternateView.CreateAlternateViewFromString("", System.Text.Encoding.UTF8, "text/plain")
                    Mail.AlternateViews.Add(alt)

                    Dim arrImagens As New ArrayList
                    Dim listaImagens As String = "|"

                    For Each src As Match In Regex.Matches(Corpo, "url\(['|\""]+.*['|\""]\)|src=[""|'][^""']+[""|']", RegexOptions.IgnoreCase)
                        If InStr(1, listaImagens, "|" & src.Value & "|") = 0 Then
                            arrImagens.Add(src.Value)
                            listaImagens &= src.Value & "|"
                        End If
                    Next

                    CIDS.Clear()

                    For indx As Integer = 0 To arrImagens.Count - 1
                        Dim cid As String = "cid:EmbedRes_" & indx + 1
                        Corpo = Corpo.Replace(arrImagens(indx), "src=""" & cid & """")
                        Dim img As String = Regex.Replace(arrImagens(indx), "url\(['|\""]", "")
                        img = Regex.Replace(img, "src=['|\""]", "")
                        img = Regex.Replace(img, "['|\""]\)", "").Replace("""", "")


                        ' redirecionamentos
                        If Not IsNothing(UrlsLocais) Then
                            For Z = 0 To UrlsLocais.Count - 1 Step 2
                                Dim urlcomp As String = UrlsLocais(Z)
                                If img.StartsWith(urlcomp, StringComparison.OrdinalIgnoreCase) Then
                                    img = img.Replace(urlcomp, UrlsLocais(Z + 1))
                                End If
                            Next
                        End If

                        Dim URL As New System.Uri(img)
                        If URL.Scheme = "http" Or URL.Scheme = "ftp" Then
                            ' carrega imagens caso remotas

                            Try
                                Dim request As HttpWebRequest = WebRequest.Create(URL)

                                request.Timeout = 5000 ' cinco segundo de carga, senão erro...
                                Dim response As HttpWebResponse = request.GetResponse()
                                Dim bmp As New Bitmap(response.GetResponseStream)

                                Dim DirTemp As String = NZV(WebConf("dir_temp"), "")
                                If DirTemp <> "" Then
                                    img = NomeArqLivre(DirTemp, "EnviaEmail")
                                Else
                                    img = System.IO.Path.GetTempFileName()
                                End If

                                If Not TMPS.Contains(img) Then
                                    TMPS.Add(img)
                                End If
                                bmp.Save(img)
                            Catch EX As Exception
                                Throw New Exception(EX.Message & " ao tentar obter conteúdo """ & URL.AbsolutePath & """")
                            End Try



                        End If
                        CIDS.Add(img)

                    Next

                    ' incorpora imagens
                    alt = AlternateView.CreateAlternateViewFromString(Corpo, System.Text.Encoding.UTF8, "text/html")

                    For z = 0 To CIDS.Count - 1
                        Dim res As New LinkedResource(CType(CIDS(z), String))
                        res.ContentId = "EmbedRes_" & z + 1
                        alt.LinkedResources.Add(res)
                    Next
                    Mail.AlternateViews.Add(alt)
                End If
            End If

            ' inclui attachados
            If Not IsNothing(Attachs) Then
                For Each attach As Object In Attachs
                    Try

                        If TypeOf attach Is String AndAlso attach <> "" Then
                            Mail.Attachments.Add(New Attachment(FileExpr(attach)))
                        ElseIf TypeOf attach Is ListItem Then
                            Dim IT As ListItem = attach
                            Dim ITA As New System.Net.Mail.Attachment(FileExpr(IT.Value))
                            ITA.Name = IT.Text
                            Mail.Attachments.Add(ITA)
                        End If
                    Catch
                    End Try
                Next
            End If

            If IsNothing(Enviar) Then
                Enviar = New System.Net.Mail.SmtpClient(NZ(SmtpHost, WebConf("smtp_host")), NZV(NZ(SmtpPort, WebConf("smtp_port")), 25))
            End If

            If Not IsNothing(SMTPUsuario) Then
                Enviar.Credentials = New System.Net.NetworkCredential(SMTPUsuario, NZ(SMTPSenha, ""))
            End If



            Enviar.Timeout = 100000
            Enviar.Send(Mail)
            Return ""
        Catch ex As Exception
            Return MessageEx(ex, "Erro ao tentar enviar email")
        End Try
    End Function

    Public Function MessageEx(ByVal Ex As Exception, Optional ByVal MensagemCompl As String = "") As String

        ' mensagem padrão
        Dim Mensagem As String = Ex.Message

        If Not IsNothing(Ex.InnerException) AndAlso NZ(Ex.InnerException.Message, "") <> "" Then
            Mensagem &= ". " & Ex.InnerException.Message
        End If
        Dim Param As String

        ' mensagens específicas
        Param = RegexGroup(Mensagem, "Cannot update (.*); field not updateable", 1).Value
        If Param <> "" Then
            Mensagem = "Por restrições da base de dados, campo " & Param & " não pode ser atualizado"
        End If

        Param = RegexGroup(Mensagem, "create duplicate values in the").Value
        If Param <> "" Then
            Mensagem = "Tentativa de registro de chave duplicada"
        End If

        Param = RegexGroup(Mensagem, "Cannot set column (.*). The value violates the MaxLength.*", 1).Value
        If Param <> "" Then
            Mensagem = "Tamanho do campo " & Param & " excede o limite"
        End If

        Param = RegexGroup(Mensagem, "The path is not of a legal").Value
        If Param <> "" Then
            Mensagem = "Caminho de arquivo inexistente ou ilegal"
        End If

        Param = RegexGroup(Mensagem, "Duplicate entry (.*) for key .*", 1).Value
        If Param <> "" Then
            Mensagem = "Tentativa de gravação de registro duplicado - " & Param
        End If

        Param = RegexGroup(Mensagem, "Empty path name is not legal").Value
        If Param <> "" Then
            Mensagem = "Nome de arquivo incorreto"
        End If

        Param = RegexGroup(Mensagem, "Could not find file '(.*?)'", 1).Value
        If Param <> "" Then
            Mensagem = "Arquivo não encontrado: " & Param
        End If

        Param = RegexGroup(Mensagem, "Thread was being aborted|O thread estava sendo anulado").Value
        If Param <> "" Then
            Mensagem = "É necessário logar-se ou sua sessão foi encerrada."
        End If


        ' ------------------------------------------------
        ' TRATAMENTO DE ERROS DO ORACLE
        If InStr(Mensagem, "ORA-01400:") <> 0 Then
            Mensagem = "Campo de identificação do registro não pode estar nulo"
        End If

        Param = RegexGroup(Mensagem, "ORA-00372:").Value
        If Param <> "" Then
            Mensagem = "Base de dados em condição de apenas para leitura ou parada para manutenção. Informe sua necessidade ao suporte"
        End If

        Param = RegexGroup(Mensagem, "ORA-02291: .*\((.*)\)", 1).Value
        If Param <> "" Then
            Mensagem = "Falta de registro relacionado em " & Param
        End If

        Param = RegexGroup(Mensagem, "ORA-00001: .*\((.*)\)", 1).Value
        If Param <> "" Then
            Mensagem = "Tentativa de registro de chave duplicada em " & Param
        End If

        Param = RegexGroup(Mensagem, "ORA-01017:").Value
        If Param <> "" Then
            Mensagem = "Logon incorreto. Usuário ou senha inválidos ou sessão expirada"
        End If

        Param = RegexGroup(Mensagem, "ORA-00942:").Value
        If Param <> "" Then
            Mensagem = "Tabela ou visão inexistente"
        End If

        Param = RegexGroup(Mensagem, "ORA-12541:|ORA-12170:").Value
        If Param <> "" Then
            Mensagem = "Banco de dados indisponível no momento. Suporte já foi contactado"
        End If

        ' ------------------------------------------------
        ' TRATAMENTO DE ERROS DO MYSQL

        Param = RegexGroup(Mensagem, "Access denied for user (.*)", 1).Value
        If Param <> "" Then
            Mensagem = "Acesso não autorizado para " & Param & ". Verifique usuário e senha e tente novamente"
        End If

        Mensagem = IIf(MensagemCompl <> "", MensagemCompl & ". ", "") & Mensagem & "."
        Return Mensagem
    End Function


    Function WebConf(ByVal param As String) As String
        If Compare(param, "SITE_DIR") Then
            Return FileExpr("~/")
        ElseIf Compare(param, "SITE_URL") Then
            Return URLExpr("~/")
        End If
        Return System.Configuration.ConfigurationManager.AppSettings(param)
    End Function

    Function Compare(ByVal Param1 As Object, ByVal Param2 As Object, Optional ByVal IgnoreCase As Boolean = True) As Boolean
        If IsNothing(Param1) And IsNothing(Param2) Then
            Return True
        ElseIf IsNothing(Param1) Or IsNothing(Param2) Then
            Return False
        Else
            If Param1.GetType.ToString = Param2.GetType.ToString Then
                If Param1.GetType.ToString = "System.String" Then
                    Return String.Compare(Param1, Param2, IgnoreCase) = 0
                Else
                    Err.Raise(20000, "IcraftBase", "Compare com tipo não previsto " & Param1.GetType.ToString & ".")
                End If
            End If
        End If
        Return False
    End Function

    Function URLExpr(ByVal ParamArray Segmentos() As Object) As String
        Dim URL As String = ExprExpr("/", "\", "", Segmentos)
        If Regex.Match(URL, "(?is)^[a-z0-9]:/").Success Then
            If Ambiente() = AmbienteTipo.WEB Then
                URL = URL.ToLower.Replace(HttpContext.Current.Server.MapPath("~/").Replace("\", "/").ToLower, "~/")
            Else
                URL = URL.Replace("\", "/").ToLower
                URL = URL.Replace(FileExpr("~/").Replace("\", "/").ToLower, "~/")
            End If
        End If
        Return URL
    End Function

    Public Enum AmbienteTipo
        Windowsforms
        WEB
    End Enum

    Public Function Ambiente() As AmbienteTipo
        Return AmbienteTipo.Windowsforms
    End Function


    Public Function TermosStrToLista(ByVal Email As Object) As ArrayList
        If TypeOf (Email) Is ArrayList Then
            Email = Join(CType(Email, ArrayList).ToArray, ";")
        End If
        Dim Lista As New ArrayList
        If NZ(Email, "") <> "" Then
            Dim Emails As Array = Split(Join(Split(Email, vbCrLf), ";"), ";")
            For Each Item As String In Emails
                Item = Trim(Item)
                If Item <> "" Then
                    Dim pref As String
                    If Item.StartsWith("bcc:", StringComparison.OrdinalIgnoreCase) Then
                        pref = "bcc:"
                        Item = Item.Substring(4)
                    Else
                        pref = ""
                    End If
                    If Item.StartsWith("conf.", StringComparison.OrdinalIgnoreCase) Then
                        Dim Result As ArrayList = TermosStrToLista(WebConf(Item.Substring(5)))
                        If pref <> "" Then
                            For Each ResultItem As String In Result
                                Lista.Add(pref & ResultItem)
                            Next
                        Else
                            Lista.AddRange(Result)
                        End If
                    Else
                        Lista.Add(pref & Item)
                    End If
                End If
            Next
        End If
        Return Lista
    End Function


    Public Function EmailStr(ByVal Email As String) As String
        Email = Trim(Email)
        Email = Email.Replace("[", "<").Replace("]", ">").Replace(Chr(160), " ")
        If Email.StartsWith("'") Then
            Email = Regex.Replace(Email, "'(.*)'", """$1""")
        End If

        Email = Email.Replace("'", "`")

        Dim SoEmail As String = ""
        If Email.IndexOf("<") = -1 Then
            SoEmail = SoEmailStr(Email)
            If SoEmail <> "" Then
                Email = ReplRepl(Email, SoEmail, "")
            End If
        Else
            SoEmail = Regex.Match(Email, "<(.*?)>").Groups(1).Value
            Email = ReplRepl(Email, "<" & SoEmail & ">", "")
        End If

        Email = ReplRepl(Email, Chr(9), "")
        Email = TrimCarac(Trim(ReplRepl(Email, "  ", " ")), New String() {Chr(34), "'"})
        Email = Regex.Replace(Email, "`(.*)`", "$1")

        If Email <> "" Then
            Email = SqlExpr(Email, """")
        End If
        Email = ExprExpr(" ", "", Email, "<" & SoEmail & ">")
        Return Email
    End Function


    Function ExprExpr(ByVal Delim As String, ByVal DelimAlternativo As String, ByVal Inicial As Object, ByVal ParamArray Segmentos() As Object) As String
        Inicial = NZ(Inicial, "")
        Dim Lista As ArrayList = ParamArrayToArrayList(Segmentos)
        For Each item As Object In Lista
            If Not IsNothing(item) Then
                If Not IsNothing(DelimAlternativo) AndAlso DelimAlternativo <> "" Then
                    item = item.Replace(DelimAlternativo, Delim)
                End If
                item = NZ(item, "")
                If item <> "" Then
                    If Inicial <> "" Then
                        If Inicial.EndsWith(Delim) AndAlso item.StartsWith(Delim) Then
                            Inicial &= CType(item, String).Substring(Delim.Length)
                        ElseIf Inicial.EndsWith(Delim) OrElse item.StartsWith(Delim) Then
                            Inicial &= item
                        Else
                            Inicial &= Delim & item
                        End If
                    Else
                        Inicial &= item
                    End If
                End If
            End If
        Next
        Return Inicial
    End Function


    Public Class Email
        Private _completo As String = ""
        Private _soendereco As String = ""
        Private _descricao As String = ""
        Private _dominio As String = ""
        Private _primeironome As String = ""
        Private _ultimonome As String = ""

        ''' <summary>
        ''' Verifica a existência de caracteres inválidos para o formato de email padrão.
        ''' </summary>
        ''' <param name="Email">Email a ser verificado.</param>
        ''' <value>Endereço de email.</value>
        ''' <returns>True se email é válido ou false caso contrário.</returns>
        ''' <remarks></remarks>
        ReadOnly Property Valida(ByVal Email As String) As Boolean
            Get
                Return Regex.IsMatch(EmailStr(Email), "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))")
            End Get
        End Property

        ''' <summary>
        ''' Verifica email carregado anteriormente.
        ''' </summary>
        ''' <value>True se email é válido ou false caso contrário.</value>
        ''' <returns>True se email é válido ou false caso contrário.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Valida() As Boolean
            Get
                Return Valida(_completo)
            End Get
        End Property

        ''' <summary>
        ''' Decompõe email em elementos.
        ''' </summary>
        ''' <param name="Email">Endereço de email.</param>
        ''' <remarks></remarks>
        Sub New(ByVal Email As String)
            _completo = EmailStr(Email)
            _soendereco = SoEmailStr(_completo)
            _descricao = Trim(RegexGroup(_completo, "\""(.*)\""", 1).Value)
            _dominio = RegexGroup(_soendereco, "@(.*)$", 1).Value

            Dim ems() As String = Split(_descricao & " ", " ")
            _primeironome = Trim(ems(0))
            _ultimonome = Trim(ems(ems.Length - 2))
        End Sub

        ''' <summary>
        ''' Domínio do email.
        ''' </summary>
        ''' <value>Domínio do email (depois do arroba).</value>
        ''' <returns>Domínio do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Dominio() As String
            Get
                Return _dominio
            End Get
        End Property

        ''' <summary>
        ''' Email completo já formatado.
        ''' </summary>
        ''' <value>Email completo já formatado.</value>
        ''' <returns>Email completo já formatado.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Completo() As String
            Get
                Return _completo
            End Get
        End Property

        ''' <summary>
        ''' Só o endereço do email (antes do arroba).
        ''' </summary>
        ''' <value>Só o endereço do email.</value>
        ''' <returns>Só o endereço do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SoEndereco() As String
            Get
                Return _soendereco
            End Get
        End Property

        ''' <summary>
        ''' Descrição do email (trecho entre apóstrofos antes do email).
        ''' </summary>
        ''' <value>Descrição do email (trecho entre apóstrofos antes do email).</value>
        ''' <returns>Descrição do email (trecho entre apóstrofos antes do email).</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Descricao() As String
            Get
                Return _descricao
            End Get
        End Property

        ''' <summary>
        ''' Primeiro nome na descrição do email.
        ''' </summary>
        ''' <value>Primeiro nome na descrição do email.</value>
        ''' <returns>Primeiro nome na descrição do email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PrimeiroNome() As String
            Get
                Return _primeironome
            End Get
        End Property

        ''' <summary>
        ''' Último nome na descrição de email.
        ''' </summary>
        ''' <value>Último nome na descrição de email.</value>
        ''' <returns>Último nome na descrição de email.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property UltimoNome() As String
            Get
                Return _ultimonome
            End Get
        End Property
    End Class


    Function NZV(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim Result As Object = NZ(Valor, Def)
        If TypeOf Result Is String AndAlso Result = "" Then
            Return Def
        ElseIf TypeOf Result Is Decimal AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Double AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Single AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Int32 AndAlso Result = 0 Then
            Return Def
        ElseIf TypeOf Result Is Byte AndAlso Result = 0 Then
            Return Def
        End If
        Return Result
    End Function


    Function NZ(ByVal Valor As Object, Optional ByVal Def As Object = Nothing) As Object
        Dim tipo As String

        If Not IsNothing(Def) Then
            tipo = Def.GetType.ToString
        ElseIf IsNothing(Valor) Then
            Return Nothing
        Else
            tipo = Valor.GetType.ToString.Trim
        End If

        If IsNothing(Valor) OrElse IsDBNull(Valor) OrElse ((tipo = "System.DateTime" Or Valor.GetType.ToString = "System.DateTime") AndAlso Valor = CDate(Nothing)) Then
            Valor = Def
        End If

        Select Case tipo
            Case "System.Decimal"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Decimal)
                End If
                Return CType(Valor, Decimal)
            Case "System.String"
                If Valor.GetType.ToString = "System.Byte[]" Then
                    Return CType(ByteArrayToObject(Valor), String)
                End If
                If Valor.GetType.IsEnum Then
                    Return Valor.ToString
                End If
                Return CType(Valor, String)
            Case "System.Double"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Double)
                End If
                Return CType(Valor, Double)
            Case "System.Boolean"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return False
                End If
                Return CType(Valor, Boolean)
            Case "System.DateTime"
                Return CType(Valor, System.DateTime)
            Case "System.Single"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Single)
                End If
                Return CType(Valor, System.Single)
            Case "System.Byte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Byte)
                End If
                Return CType(Valor, System.Byte)
            Case "System.Char"
                Return CType(Valor, System.Char)
            Case "System.SByte"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, SByte)
                End If
                Return CType(Valor, System.SByte)
            Case "System.Int32"
                If Valor.GetType.ToString = "System.String" AndAlso Valor = "" Then
                    Return CType(0, Int32)
                End If
                Return CType(Valor, Int32)
            Case "System.DBNull"
                Return Valor
            Case "System.Collections.ArrayList"
                Return ParamArrayToArrayList(Valor)
            Case "System.Data.DataSet"
                If IsNothing(Valor) Then
                    Return Def
                End If
                Return Valor
        End Select

        Return CType(Valor, String)
    End Function

    Function ParamArrayToArrayList(ByVal ParamArray Params() As Object) As Object

        ' caso não existam parâmetros
        If IsNothing(Params) OrElse Params.Length = 0 Then
            Return New ArrayList
        End If

        ' caso já seja um arraylist
        If Params.Length = 1 And TypeOf (Params(0)) Is ArrayList Then
            Return Params(0)
        End If

        ' caso tenha que juntar
        Dim ListaParametros As ArrayList = New ArrayList
        For Each Item As Object In Params
            If Not IsNothing(Item) Then

                ' >> TIPOS PREVISTOS EM ARRAYLIST...
                ' array
                ' arraylist
                ' string
                ' dataset
                ' datarowcollection

                If TypeOf Item Is Array Then
                    For Each SubItem As Object In Item
                        ListaParametros.AddRange(ParamArrayToArrayList(SubItem))
                    Next
                ElseIf TypeOf Item Is ArrayList OrElse Item.GetType.ToString.StartsWith("System.Collections.Generic.List") Then
                    ListaParametros.AddRange(Item)
                ElseIf TypeOf Item Is String Then
                    ListaParametros.Add(Item)
                ElseIf TypeOf Item Is DataSet Then
                    For Each Row As DataRow In Item.Tables(0).rows
                        For Each Campo As Object In Row.ItemArray
                            ListaParametros.Add(Campo)
                        Next
                    Next
                ElseIf TypeOf Item Is DataRow Then
                    For Each Campo As Object In CType(Item, DataRow).ItemArray
                        ListaParametros.Add(Campo)
                    Next
                ElseIf TypeOf Item Is System.IO.FileInfo Then
                    ListaParametros.Add(Item.name)
                Else
                    ListaParametros.Add(Item)
                End If
            End If
        Next
        Return ListaParametros
    End Function


    Function NomeArqLivre(ByVal NomeDir As String, ByVal NomeArq As String) As String
        Dim DD As New System.IO.DirectoryInfo(NomeDir)
        If Not DD.Exists Then
            DD.Create()
            NomeArq = FileExpr(NomeDir, NomeArq)
        Else
            Dim z As Integer = 1
            Dim NomeTest As String = FileExpr(NomeDir, NomeArq)
            Do While True
                If z <> 1 Then
                    NomeTest = FileExpr(NomeDir, System.IO.Path.GetFileNameWithoutExtension(NomeArq) & "_" & Trim(Format(z, "    00")) & System.IO.Path.GetExtension(NomeArq))
                End If
                Dim FF As New System.IO.FileInfo(NomeTest)
                If Not FF.Exists Then
                    Exit Do
                End If
                z += 1
            Loop
            NomeArq = NomeTest
        End If
        Return NomeArq
    End Function

    Function FileExpr(ByVal ParamArray Segmentos() As String) As String
        Dim Raiz As String = New System.Web.UI.Control().ResolveUrl("~/").Replace("/", "\")
        Dim Arq As String = ExprExpr("\", "/", "", Segmentos)
        If Arq.StartsWith(Raiz) Then
            Arq = "~\" & Mid(Arq, Len(Raiz) + 1)
        End If

        If Arq.StartsWith("~\") Then
            If Ambiente() = AmbienteTipo.WEB Then
                Arq = HttpContext.Current.Server.MapPath(Arq)
            Else
                Dim DirExec As String = FileExpr(WebConf("dir_raiz_site"), "\")
                If DirExec = "" Or DirExec = "\" Then
                    DirExec = System.Windows.Forms.Application.ExecutablePath
                End If
                Arq = Arq.Replace("~\", System.IO.Path.GetDirectoryName(DirExec) & "\")
            End If
        End If
        Return Arq
    End Function

    Function ByteArrayToObject(ByVal Bytes() As Byte) As Object
        Dim Obj As Object = Nothing
        Try
            Dim fs As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim formatter As System.Runtime.Serialization.Formatters.Binary.BinaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            fs.Write(Bytes, 0, Bytes.Length)
            fs.Seek(0, IO.SeekOrigin.Begin)

            Obj = formatter.Deserialize(fs)
        Catch
        End Try
        Return Obj
    End Function

    Public Function SoEmailStr(ByVal Email As String) As String
        Return RegexGroup(Email, "(^|[ \t\[\<\>\""]*)([\w-.]+@[\w-]+(\.[\w-]+)+)(($|[ \t\<\>\""]*))", 2).Value
    End Function


    Function RegexGroup(ByVal Texto As String, ByVal Mascara As String, Optional ByVal Grupo As Object = 0) As System.Text.RegularExpressions.Group
        Return System.Text.RegularExpressions.Regex.Match(NZ(Texto, ""), Mascara).Groups(Grupo)
    End Function

    Function ReplRepl(ByVal Texto As String, ByVal De As String, ByVal Para As String) As String
        Do While InStr(Texto, De) <> 0
            Texto = Replace(Texto, De, Para)
        Loop
        Return Texto
    End Function

    Function SqlExpr(ByVal Conteudo As Object, Optional ByVal CaracAbreFechaString As String = "'") As String
        If TypeOf (Conteudo) Is String Then
            Return CaracAbreFechaString & Replace(Conteudo, CaracAbreFechaString, CaracAbreFechaString & CaracAbreFechaString) & CaracAbreFechaString
        ElseIf TypeOf (Conteudo) Is DBNull Then
            Return "NULL"
        ElseIf TypeOf Conteudo Is Decimal OrElse TypeOf Conteudo Is Double OrElse TypeOf Conteudo Is Single OrElse TypeOf Conteudo Is Int32 OrElse TypeOf Conteudo Is Byte Then
            Return Str(Conteudo)
        ElseIf TypeOf (Conteudo) Is Boolean Then
            Return IIf(Conteudo, Boolean.TrueString, Boolean.FalseString)
        ElseIf TypeOf (Conteudo) Is Date Then
            Return "#" & Format(Conteudo, "yyyy-MM-dd HH:mm:ss") & "#"
        Else
            Throw New Exception("Tipo desconhecido " & Conteudo.GetType.ToString & " para obtenção de expressão para sql.")
        End If
    End Function

    Public Function TrimCarac(ByVal Texto As String, ByVal Carac() As String) As String
        Dim Achou As Boolean = True
        Do While Achou
            Achou = False
            For Each Item As String In Carac
                Do While Texto.StartsWith(Item, StringComparison.OrdinalIgnoreCase)
                    Texto = Mid(Texto, Len(Item) + 1)
                    Achou = True
                Loop
                Do While Texto.EndsWith(Item, StringComparison.OrdinalIgnoreCase)
                    Texto = StrStr(Texto, 0, -Len(Item))
                Loop
            Next
        Loop
        Return Texto
    End Function

    Function StrStr(ByVal Variavel As String, ByVal Inicio As Integer, Optional ByVal Final As Integer = Nothing) As String
        If Inicio < 0 Then
            Inicio = (Len(Variavel) + Inicio)
        End If
        If Not NZ(Final, 0) = 0 Then
            If Final < 0 Then
                Final = (Len(Variavel) + Final) - 1
            End If
            Return Variavel.Substring(Inicio, Final - Inicio + 1)
        End If
        Return Variavel.Substring(Inicio)
    End Function

    Function Entifica(ByVal Texto As String, Optional ByVal Tipo As TipoEntifica = TipoEntifica.Tudo) As String
        Dim G1() As String = {"&", "¡", "¢", "£", "¤", "¥", "¦", "§", "¨", "©", "ª", "«", "¬", "¬", "®", "¯", "°", "±", "²", "³", "´", "µ", "¶", "•", "¸", "¹", "º", "»", "¼", "½", "¾", "¿", "×", "÷", "Æ", "Ð", "Ø", "Þ", "ß", "æ", "ø", "þ"}
        Dim G2() As String = {"À", "Á", "Â", "Ã", "Ä", "Å", "Ç", "È", "É", "Ê", "Ë", "Ì", "Í", "Î", "Ï", "Ñ", "Ò", "Ó", "Ô", "Õ", "Ö", "Ù", "Ú", "Û", "Ü", "Ý", "à", "á", "â", "ã", "ä", "å", "ç", "è", "é", "ê", "ë", "ì", "í", "î", "ï", "ð", "ñ", "ò", "ó", "ô", "õ", "ö", "ù", "ú", "û", "ü", "ý", "ÿ"}
        Dim G3() As String = {"""", "'", "<", ">"}

        If Tipo = TipoEntifica.Tudo Then
            For Each IT As String In G1
                Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
            Next
        End If

        If Tipo = TipoEntifica.Tudo Or Tipo = TipoEntifica.SoAcentos Then
            For Each IT As String In G2
                Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
            Next
        End If

        If Tipo = TipoEntifica.Tudo And Not Tipo = TipoEntifica.MenosHTML Then
            For Each IT As String In G3
                Texto = Texto.Replace(IT, "&#" & Asc(IT) & ";")
            Next
        End If

        Return Texto
    End Function

    Public Enum TipoEntifica
        Tudo
        SoAcentos
        MenosHTML
    End Enum


End Module
