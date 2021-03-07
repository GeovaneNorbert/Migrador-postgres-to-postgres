
Imports Npgsql
Imports System.Data
Imports System.Security
Imports System.IO
Imports System.Text
Public Class Form1
    Dim conn1 As NpgsqlConnection
    Dim conn2 As NpgsqlConnection
    Dim chaveErro As String
    Dim Conexao1 As String
    Dim Conexao2 As String
    Dim StrLog As String
    Dim NoLinhas As String
    Dim msg As MessageBox
    Dim NoRepeticoes As Integer
    Dim StrEmpresa1 As String
    Dim StrEmpresa2 As String
    Dim cronometro As New Stopwatch


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Servidor1.Text = "localhost"
        Senha1.PasswordChar = "*"
        Usuario1.Text = "postgres"
        Senha1.Text = "postgres"
        database1.Text = "darios"
        Empresa1.Text = "ep001"

        servidor2.Text = "localhost"
        Senha2.PasswordChar = "*"
        Usuario2.Text = "postgres"
        Senha2.Text = "postgres"
        Database2.Text = "darios"
        Empresa2.Text = "ep001"

        Linhas.Text = "1"
        Repeticoes.Text = "10"
        chaveErro = "0"
        Label8.Text = "0"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Conexao1 = "Server=" + Servidor1.Text + ";Port=5432;UserId=" + Usuario1.Text + ";Password=" + Senha1.Text + ";Database=" + database1.Text & ";Timeout=1024;CommandTimeout=3600;"

        If CheckBox1.Checked Then
            Conexao2 = "Server=" + Servidor1.Text + ";Port=5432;UserId=" + Usuario1.Text + ";Password=" + Senha1.Text + ";Database=" + database1.Text & ";Timeout=1024;CommandTimeout=3600;"
        Else
            Conexao2 = "Server=" + servidor2.Text + ";Port=5432;UserId=" + Usuario2.Text + ";Password=" + Senha2.Text + ";Database=" + Database2.Text & ";Timeout=1024;CommandTimeout=3600;"
        End If

        StrEmpresa1 = Empresa1.Text
        StrEmpresa2 = Empresa2.Text

        conn1 = New NpgsqlConnection(Conexao1)
        conn2 = New NpgsqlConnection(Conexao2)
        Try
            conn1.Open()
            If (conn1.State = ConnectionState.Open) Then
                MessageBox.Show(“Conexão com origem bem sucedida!”)
            Else
                MessageBox.Show("nao abriu origem")
            End If
        Catch ex As Exception
            MessageBox.Show("falha na conexao com origem")
        End Try
        Try
            conn2.Open()
            If (conn2.State = ConnectionState.Open) Then
                MessageBox.Show(“Conexão com destino bem sucedida!”)
            Else
                MessageBox.Show("nao abriu destino")
            End If
        Catch ex As Exception
            MessageBox.Show("falha na conexao com destino")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Iniciar.Click

        '    Label8.Text = DateTime.Now.ToString()
        '    Label11.Text = "0"
        '    NoRepeticoes = Repeticoes.Text
        '    NoLinhas = Linhas.Text
        '    cronometro.Start()
        '    Dim i As Integer
        '    i = 1
        '    While i <= NoRepeticoes
        '        If i > 1 Then
        '            Dim myProcess As New Process
        '            Label11.Text = cronometro.Elapsed.ToString
        '            Dim result As DialogResult = MessageBox.Show("Deseja continuar", "Deletando", MessageBoxButtons.YesNoCancel)
        '            If result = DialogResult.Cancel Then
        '                Exit While
        '            ElseIf result = DialogResult.No Then
        '                Exit While
        '            ElseIf result = DialogResult.Yes Then

        '            End If
        '        End If
        '        deleta()
        '        i = i + 1
        '    End While
        '    StrLog = StrLog & " CONCLUIDO " & vbCrLf
        '    Log.Text = StrLog
        '    Log.Refresh()
        '    cronometro.Stop()
        '    Label11.Text = cronometro.Elapsed.ToString
        '    cronometro.Reset()
        'End Sub

        'Private Sub deleta()
        '    Dim Sql As String = "begin; "
        '    Sql = Sql & " delete FROM " & StrEmpresa1 & ".es_tbsaldos  "
        '    Sql = Sql & "using " & StrEmpresa1 & ".es_tbsaldos tb  "
        '    Sql = Sql & "left join  "
        '    Sql = Sql & "(  "
        '    Sql = Sql & "select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa1 & ".Pd_TblItens "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa1 & ".PR_ITPEDIDO "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "select produt_ifa, sequen_ifa from ep001.pd_itfabric "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa1 & ".TITENTRA) as vendas "
        '    Sql = Sql & "on produt_pdi=tb.codigo_sld and sgrupo_pdi=tb.sequen_sld "
        '    Sql = Sql & "where produt_pdi is null "
        '    Sql = Sql & "and tb.mchave_sld=" & StrEmpresa1 & ".es_tbsaldos.mchave_sld "
        '    Sql = Sql & "and tb.mchave_sld >" & chaveErro
        '    Sql = Sql & "and tb.mchave_sld<= "
        '    Sql = Sql & "(select max(mchave_sld) from ( "
        '    Sql = Sql & "select mchave_sld FROM " & StrEmpresa1 & ".es_tbsaldos tb "
        '    Sql = Sql & "left join "
        '    Sql = Sql & "( "
        '    Sql = Sql & "select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa1 & ".Pd_TblItens "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa1 & ".PR_ITPEDIDO "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "select produt_ifa, sequen_ifa from ep001.pd_itfabric "
        '    Sql = Sql & "UNION "
        '    Sql = Sql & "SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa1 & ".TITENTRA) as vendas "
        '    Sql = Sql & "on produt_pdi=tb.codigo_sld and sgrupo_pdi=tb.sequen_sld "
        '    Sql = Sql & "where produt_pdi is null "
        '    Sql = Sql & "and mchave_sld >" & chaveErro
        '    Sql = Sql & "order by mchave_sld "
        '    Sql = Sql & "limit " & NoLinhas & " ) as a); commit; "

        '    Try
        '        If (conn1.State = ConnectionState.Closed) Then
        '            conn1.Open()
        '        End If

        '        Dim Command As NpgsqlCommand = New NpgsqlCommand(Sql, conn1)
        '        Command.ExecuteNonQuery()
        '        StrLog = StrLog & "" & NoLinhas & " LINHAS EXECUTADAS " & vbCrLf
        '        Log.Text = StrLog
        '        Log.Refresh()


        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '        chaveErro = (PegaMaiorChave())
        '        StrLog = StrLog & "Erro ao executar o bloco de " & NoLinhas & " linhas onde o maior registro é " & chaveErro & "" & vbCrLf
        '        Log.Text = StrLog
        '        Log.Refresh()
        '    Finally

        '    End Try
    End Sub

    Public Function PegaMaiorChave() As String
        Dim ConsultaMaiorChave As NpgsqlConnection
        Dim ds As New DataSet()
        Dim da As NpgsqlDataAdapter
        Dim StringSQL As String
        ConsultaMaiorChave = New NpgsqlConnection(Conexao1)
        Try
            ConsultaMaiorChave.Open()

            StringSQL = "select max(mchave_sld) from ( "
            StringSQL = StringSQL & " select mchave_sld FROM " & StrEmpresa1 & ".es_tbsaldos tb "
            StringSQL = StringSQL & " left join "
            StringSQL = StringSQL & "( "
            StringSQL = StringSQL & " select PRODUT_pdi,sgrupo_pdI from " & StrEmpresa1 & ".Pd_TblItens "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " SELECT PRODUT_PRI,SGRUPO_PRI FROM " & StrEmpresa1 & ".PR_ITPEDIDO "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " select produt_ifa, sequen_ifa from ep001.pd_itfabric "
            StringSQL = StringSQL & " UNION "
            StringSQL = StringSQL & " SELECT PRODUT_ENT,SEQUEN_ENT FROM " & StrEmpresa1 & ".TITENTRA) as vendas "
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
        If (conn1.State = ConnectionState.Open) Then
            conn1.Close()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            servidor2.Enabled = False
            Usuario2.Enabled = False
            Senha2.Enabled = False
            Database2.Enabled = False
        Else
            servidor2.Enabled = True
            Usuario2.Enabled = True
            Senha2.Enabled = True
            Database2.Enabled = True
        End If
    End Sub

    Private Sub Listar_Click(sender As Object, e As EventArgs) Handles Listar.Click
        ListBox1.DataSource = carregaTabelasComDados.DefaultView
        ListBox1.DisplayMember = "relname"
        ListBox1.ValueMember = "relname"
        ListBox1.Refresh()
    End Sub
    Private Function carregaTabelasComDados() As DataTable
        Dim BuscaTabelas As NpgsqlConnection
        Dim TT As New DataTable
        Dim da As NpgsqlDataAdapter
        Dim StringSQL As String
        BuscaTabelas = New NpgsqlConnection(Conexao1)
        Try
            BuscaTabelas.Open()

            StringSQL = " SELECT relname "
            StringSQL = StringSQL & " FROM pg_stat_user_tables "
            StringSQL = StringSQL & " where schemaname = '" & StrEmpresa1 & "'"
            StringSQL = StringSQL & " And n_live_tup > 0"
            StringSQL = StringSQL & " ORDER BY relname;"

            da = New NpgsqlDataAdapter(StringSQL, BuscaTabelas)

            da.Fill(TT)

            Return TT

        Catch ex As Exception

            MsgBox("abra a conexao primeiro")

            '  Return 0
            Return TT
        Finally
            If BuscaTabelas.State = ConnectionState.Open Then
                BuscaTabelas.Close()
            End If
        End Try
    End Function

    Private Sub Transferir_Click(sender As Object, e As EventArgs) Handles Transferir.Click
        Dim BuscaColunas As NpgsqlConnection
        Dim ds As New DataSet()
        Dim da As NpgsqlDataAdapter
        Dim StringSQL As String
        Dim StringColunas As String
        Dim StringMerge As String
        Dim row As DataRow
        Dim nomeTabela As String

        For Each SelectedItem As DataRowView In ListBox1.SelectedItems
            row = SelectedItem.Row
            nomeTabela = row(0).ToString()


            BuscaColunas = New NpgsqlConnection(Conexao2)
            StringSQL = "Select column_name||',' column_name from information_schema.columns "
            StringSQL = StringSQL & "where table_schema ='" & StrEmpresa2 & "' and table_name='" & nomeTabela & "' "

            da = New NpgsqlDataAdapter(StringSQL, BuscaColunas)
            Dim TT As New DataTable
            da.Fill(TT)

            Dim row2 As DataRow
            For Each row2 In TT.Rows
                StringColunas = StringColunas & row2(0).ToString
            Next
            StringColunas = StringColunas.Substring(0, StringColunas.Length - 1)
            StringMerge = "INSERT INTO " & StrEmpresa2 & "." & nomeTabela & "( " & StringColunas & " )"
            StringMerge += "SELECT" & StringColunas & " FROM " & StrEmpresa1 & "." & nomeTabela
            StringMerge += " ON CONFLICT "

            MsgBox(StringColunas)
        Next


    End Sub

    'SELECT schemaname,relname,n_live_tup 
    'FROM pg_stat_user_tables 
    'where schemaname = 'ep001'-- and relname='es_tbsaldos'
    'And n_live_tup > 0
    'ORDER BY relname;

    ' Selec* From ep001.es_tbsaldos


    'Select Case column_name from information_schema.columns
    'where table_schema ='ep003' and table_name='es_tbsaldos'


    'INSERT INTO ep001.es_tbmarcas (codigo_mar,descri_mar,regfab_mar,operad_mar,dcadas_mar,filial_mar,export_drv,dalter_sis,pcomis_mar,pdesco_mar,mlucro_mar)
    '   Select Case codigo_sco, descri_sco, null, operad_sco, dcadas_sco, filial_sco,export_drv,dalter_sis,null,null,null from ep001.es_tblsecao
    'On CONFLICT (codigo_mar) DO UPDATE SET 
    '  (codigo_mar,descri_mar,regfab_mar,operad_mar,dcadas_mar,filial_mar,export_drv,dalter_sis,pcomis_mar,pdesco_mar,mlucro_mar)=(EXCLUDED.codigo_mar,EXCLUDED.descri_mar,EXCLUDED.regfab_mar,EXCLUDED.operad_mar,EXCLUDED.dcadas_mar,EXCLUDED.filial_mar,EXCLUDED.export_drv,EXCLUDED.dalter_sis,EXCLUDED.pcomis_mar,EXCLUDED.pdesco_mar,EXCLUDED.mlucro_mar);




End Class


