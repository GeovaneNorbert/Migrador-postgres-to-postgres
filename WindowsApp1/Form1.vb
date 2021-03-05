
Imports Npgsql
Imports System.Data
Imports System.Security
Imports System.IO
Imports System.Text
Public Class Form1
    Dim conn As NpgsqlConnection
    Dim chaveErro As String
    Dim Conexao As String
    Dim StrLog As String
    Dim NoLinhas As String
    Dim msg As MessageBox
    Dim NoRepeticoes As Integer
    Dim StrEmpresa As String
    Dim cronometro As New Stopwatch


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Servidor.Text = "localhost"
        Senha.PasswordChar = "*"
        Usuario.Text = "postgres"
        Senha.Text = "postgres"
        database.Text = "portas"
        Linhas.Text = "1"
        Repeticoes.Text = "10"
        chaveErro = "0"
        Empresa.Text = "ep001"
        Label8.Text = "0"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Conexao = "Server=" + Servidor.Text + ";Port=5432;UserId=" + Usuario.Text + ";Password=" + Senha.Text + ";Database=" + database.Text & ";Timeout=1024;CommandTimeout=3600;"
        StrEmpresa = Empresa.Text

        conn = New NpgsqlConnection(Conexao)
        Try
            conn.Open()
            If (conn.State = ConnectionState.Open) Then
                MessageBox.Show(“Conexão bem sucedida!”)
            Else
                MessageBox.Show("nao abriu")
            End If
        Catch ex As Exception
            MessageBox.Show("falha na conexao")

        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Iniciar.Click

        Label8.Text = DateTime.Now.ToString()
        Label11.Text = "0"
        NoRepeticoes = Repeticoes.Text
        NoLinhas = Linhas.Text
        cronometro.Start()
        Dim i As Integer
        i = 1
        While i <= NoRepeticoes
            If i > 1 Then
                Dim myProcess As New Process
                Label11.Text = cronometro.Elapsed.ToString
                Dim result As DialogResult = MessageBox.Show("Deseja continuar", "Deletando", MessageBoxButtons.YesNoCancel)
                If result = DialogResult.Cancel Then
                    Exit While
                ElseIf result = DialogResult.No Then
                    Exit While
                ElseIf result = DialogResult.Yes Then

                End If
            End If
            deleta()
            i = i + 1
        End While
        StrLog = StrLog & " CONCLUIDO " & vbCrLf
        Log.Text = StrLog
        Log.Refresh()
        cronometro.Stop()
        Label11.Text = cronometro.Elapsed.ToString
        cronometro.Reset()
    End Sub

    Private Sub deleta()
        Dim Sql As String = "begin; "
        Sql = Sql & " delete FROM " & StrEmpresa & ".es_tbsaldos  "
        Sql = Sql & "using " & StrEmpresa & ".es_tbsaldos tb  "
        Sql = Sql & "left join  "
        Sql = Sql & "(  "
        Sql = Sql & "select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa & ".Pd_TblItens "
        Sql = Sql & "UNION "
        Sql = Sql & "SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa & ".PR_ITPEDIDO "
        Sql = Sql & "UNION "
        Sql = Sql & "select produt_ifa, sequen_ifa from ep001.pd_itfabric "
        Sql = Sql & "UNION "
        Sql = Sql & "SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa & ".TITENTRA) as vendas "
        Sql = Sql & "on produt_pdi=tb.codigo_sld and sgrupo_pdi=tb.sequen_sld "
        Sql = Sql & "where produt_pdi is null "
        Sql = Sql & "and tb.mchave_sld=" & StrEmpresa & ".es_tbsaldos.mchave_sld "
        Sql = Sql & "and tb.mchave_sld >" & chaveErro
        Sql = Sql & "and tb.mchave_sld<= "
        Sql = Sql & "(select max(mchave_sld) from ( "
        Sql = Sql & "select mchave_sld FROM " & StrEmpresa & ".es_tbsaldos tb "
        Sql = Sql & "left join "
        Sql = Sql & "( "
        Sql = Sql & "select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa & ".Pd_TblItens "
        Sql = Sql & "UNION "
        Sql = Sql & "SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa & ".PR_ITPEDIDO "
        Sql = Sql & "UNION "
        Sql = Sql & "select produt_ifa, sequen_ifa from ep001.pd_itfabric "
        Sql = Sql & "UNION "
        Sql = Sql & "SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa & ".TITENTRA) as vendas "
        Sql = Sql & "on produt_pdi=tb.codigo_sld and sgrupo_pdi=tb.sequen_sld "
        Sql = Sql & "where produt_pdi is null "
        Sql = Sql & "and mchave_sld >" & chaveErro
        Sql = Sql & "order by mchave_sld "
        Sql = Sql & "limit " & NoLinhas & " ) as a); commit; "

        Try
            If (conn.State = ConnectionState.Closed) Then
                conn.Open()
            End If

            Dim Command As NpgsqlCommand = New NpgsqlCommand(Sql, conn)
            Command.ExecuteNonQuery()
            StrLog = StrLog & "" & NoLinhas & " LINHAS EXECUTADAS " & vbCrLf
            Log.Text = StrLog
            Log.Refresh()


        Catch ex As Exception
            MsgBox(ex.Message)
            chaveErro = (PegaMaiorChave())
            StrLog = StrLog & "Erro ao executar o bloco de " & NoLinhas & " linhas onde o maior registro é " & chaveErro & "" & vbCrLf
            Log.Text = StrLog
            Log.Refresh()
        Finally

        End Try
    End Sub

    Public Function PegaMaiorChave() As String
        Dim ConsultaMaiorChave As NpgsqlConnection
        Dim ds As New DataSet()
        Dim da As NpgsqlDataAdapter
        Dim StringSQL As String
        ConsultaMaiorChave = New NpgsqlConnection(Conexao)
        Try
            ConsultaMaiorChave.Open()

            StringSQL = "select max(mchave_sld) from ( "
            StringSQL = StringSQL & " select mchave_sld FROM " & StrEmpresa & ".es_tbsaldos tb "
            StringSQL = StringSQL & " left join "
            StringSQL = StringSQL & "( "
            StringSQL = StringSQL & " select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa & ".Pd_TblItens "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa & ".PR_ITPEDIDO "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " select produt_ifa, sequen_ifa from ep001.pd_itfabric "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa & ".TITENTRA) as vendas "
            StringSQL = StringSQL & " on produt_pdi=tb.codigo_sld and sgrupo_pdi=tb.sequen_sld "
            StringSQL = StringSQL & " where produt_pdi is null "
            StringSQL = StringSQL & " and mchave_sld>" & chaveErro
            StringSQL = StringSQL & " order by mchave_sld "
            StringSQL = StringSQL & " limit " & NoLinhas & " ) as a "
            da = New NpgsqlDataAdapter(StringSQL, ConsultaMaiorChave)

            da.Fill(ds)

            Return ds.Tables(0).Rows(0)(0)

        Catch ex As Exception

            MsgBox("abra a conexao primeiro")

            Return ""
        Finally
            If ConsultaMaiorChave.State = ConnectionState.Open Then
                ConsultaMaiorChave.Close()
            End If
        End Try
    End Function
    Private Sub AbreConfirmacao()

    End Sub



    Private Sub Log_TextChanged(sender As Object, e As EventArgs) Handles Log.TextChanged
        Log.SelectionStart = Log.Text.Length
        Log.ScrollToCaret()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (conn.State = ConnectionState.Open) Then
            conn.Close()
        End If
    End Sub
End Class


