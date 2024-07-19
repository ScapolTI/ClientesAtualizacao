
Imports System
Imports System.IO
Imports System.Media
Imports System.Net.Mail
Imports System.Math
Imports System.ConsoleKeyInfo
Imports System.Security.Principal.WindowsIdentity
Imports System.Data
Imports System.Data.Odbc

Public Class Form1

    Dim oCompany As New SAPbobsCOM.Company
    Dim oCompany2 As New SAPbobsCOM.Company
    Dim sErrMsg As String
    Dim sErrMsg2 As String
    Dim lErrCode As Long
    Dim lErrCode2 As Long
    Dim lConexao As Long = 1
    Dim lConexao2 As Long = 1
    Dim Tipo As String
    Dim Tipo2 As String
    Dim tSQL As String
    Dim ErrCode As String
    Dim ErrMsg As String

    Dim Usuario As String
    Dim UserSAP As String
    Dim SenhaSAP As String

    Dim cn As New ADODB.Connection()
    Dim rs As New ADODB.Recordset()
    Dim rsProcedure As New ADODB.Recordset()
    Dim cnStr As String
    Dim cmd As New ADODB.Command()


    Dim tbParceiro As SAPbobsCOM.BusinessPartners
    Dim tbContato As SAPbobsCOM.ContactEmployees

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Aguarde alguns segundos enquanto a conexão é efetuada."
        Label1.Visible = True
        Refresh()

        Usuario = GetCurrent.Name.Replace("SCAPOL\", "")

        Label3.Text = Usuario

        ' string de conexao com o banco
        cnStr = "DSN=hanab1;  UID=SYSTEM; PWD=h1n1Sc1p4l;"

        ' 2. Abre a Conexao
        cn.Open(cnStr)
        'cn.Close()

        tSQL = " Select ""U_usersap"",""U_senhasap"" from SBH_SCAPOL.""@TBUSUARIOS"" t where t.""U_userwindows"" = '" + Usuario + "' "

        cn.CommandTimeout = 360
        rs = cn.Execute(tSQL)

        If Not rs.EOF Then
            UserSAP = rs.Fields.Item("u_usersap").Value
            SenhaSAP = rs.Fields.Item("u_senhasap").Value
        Else
            MessageBox.Show("Usuario Windows não mapeado com Usuario SAP - SBH_SCAPOL.@TBUSUARIOS")
        End If

        cn.Close()

        '------------------------------------------------------
        ' Conectando no SAP
        '------------------------------------------------------


        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB

        oCompany.Disconnect()

        oCompany.Server = "NDB@hanab1:30013"
        oCompany.UseTrusted = False
        oCompany.DbUserName = "SYSTEM"
        oCompany.DbPassword = "h1n1Sc1p4l"
        oCompany.CompanyDB = "SBH_SCAPOL"
        oCompany.UserName = UserSAP
        oCompany.Password = SenhaSAP


        lConexao = oCompany.Connect

        If lConexao <> 0 Then
            oCompany.GetLastError(lErrCode, sErrMsg)
            MsgBox("Não foi possivel estabelecer a conexão!" + Chr(13) + "Por favor tentar novamente! ", MsgBoxStyle.Information, "Erro")
            Label1.Text = "Erro ao Conectar: " & sErrMsg
            Label1.Visible = True
        Else
            Label1.Text = oCompany.CompanyDB

        End If


        tbParceiro = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Button1_Click(sender, e)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim qtd As Integer
        qtd = 0

        Dim Qry As SAPbobsCOM.Recordset
        Dim Qry2 As SAPbobsCOM.Recordset
        Dim CardCode As String
        Dim Canal As String
        Dim DDD As String
        Dim Telefone As String
        Dim EmailXML As String
        Dim EmailComercial As String

        Qry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Qry2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        tSQL = ""
        tSQL = tSQL + " select ""CardCode"",""Canal"",""DDD"",""Telefone"",""EmailXML"",""EmailComercial"" "
        tSQL = tSQL + " from SBH_SCAPOL.VT_SITE_APP_CLIENTES_ATUALIZACAO "
        tSQL = tSQL + " where ""Atualiza"" = 'N' "

        Qry.DoQuery(tSQL)


        ProgressBar1.Maximum = Qry.RecordCount()
        ProgressBar1.Value = 0

        While Not Qry.EoF

            CardCode = Qry.Fields.Item("CardCode").Value.ToString
            Canal = Qry.Fields.Item("Canal").Value.ToString
            DDD = Qry.Fields.Item("DDD").Value.ToString
            Telefone = Qry.Fields.Item("Telefone").Value.ToString
            EmailXML = Qry.Fields.Item("EmailXML").Value.ToString
            EmailComercial = Qry.Fields.Item("EmailComercial").Value.ToString

            tbParceiro.GetByKey(CardCode)

            If Canal <> "" Then
                tbParceiro.UserFields.Fields.Item("U_Canal").Value = Canal
            End If

            If DDD <> "" Then
                tbParceiro.Phone2 = DDD
            End If

            If Telefone <> "" Then
                tbParceiro.Phone1 = Telefone
            End If

            If EmailXML <> "" Then
                tbParceiro.EmailAddress = EmailXML
            End If

            If EmailComercial <> "" Then
                tbParceiro.Website = EmailComercial
            End If

            If tbParceiro.Update() <> 0 Then
                oCompany2.GetLastError(ErrCode, ErrMsg)
                MsgBox(ErrCode & " " & ErrMsg)
            Else
                qtd = qtd + 1
                ProgressBar1.Value = ProgressBar1.Value + 1
            End If

            tSQL = ""
            tSQL = tSQL + " UPDATE SBH_SCAPOL.VT_SITE_APP_CLIENTES_ATUALIZACAO SET ""Atualiza"" = 'S' WHERE ""CardCode"" = '" & CardCode & "' "

            Qry2.DoQuery(tSQL)

            Qry.MoveNext()

        End While

        Label3.Visible = True
        Label3.Text = qtd.ToString() + " Clientes Atualizados"

        Application.Exit()


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Button1_Click(sender, e)
    End Sub
End Class
