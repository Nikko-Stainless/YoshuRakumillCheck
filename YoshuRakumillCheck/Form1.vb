Imports System.IO
Imports Oracle.ManagedDataAccess.Client
Imports Nikko.Windows
Imports OpenPop.Pop3

Public Class Form1
    Dim folderPath As String = ""
    Dim connectionStringForYoshu As String = ""
    Dim connectionStringForNikko As String = ""
    Dim fileINI = ""
    Private Const HOST_NAME = "27.34.144.88"
    Private Const PORT = 110
    Private MAIL_ADDRESS = "milko@nikko-sus.co.jp"
    Private MAIL_PASSWORD = "nikko-sus2"
    Private TEXT_MEMO1 = "okuisakicd||a.atukaitencd||a.todokesakicd=b.tokuisakicd||b.atukaitencd||b.todokesakicd) and a.tokuisakicd<>'99999' and a.bikou='TPA';"


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'load connect Nikko
        Dim aaaa = New AAAA()
        connectionStringForNikko = aaaa.GetConnectionString
        'load connect Yoshu
        If Not LoadFileini() Then Me.Close()
        Label1.Text = ""
        RichTextBox1.Text = TEXT_MEMO1
        TabControl1.SelectedIndex = 3
        SelectedIndexChanged()
    End Sub
    'load file ini 
    Private Function LoadFileini() As Boolean
        fileINI = Application.StartupPath + "\YoshuRakumillCheck_AppSettings.ini"

        If Not IO.File.Exists(fileINI) Then Return False

        Dim lines As String() = IO.File.ReadAllLines(fileINI)
        Dim section As String = ""

        For Each line As String In lines
            line = line.Trim()
            If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                section = line.Trim("["c, "]"c)
            ElseIf Not String.IsNullOrEmpty(line) Then
                Select Case section
                    Case "ConnectString"
                        connectionStringForYoshu = line.Trim()
                    Case "RakumillCheckFolder"
                        folderPath = line.Trim()
                End Select
            End If
        Next
        Return True
    End Function
    'SQL
    Private Function GetData(query As String, isConnectString As String) As DataTable
        Dim dataTable As New DataTable()

        Using connection As New OracleConnection(isConnectString)
            Using command As New OracleCommand(query, connection)
                Using adapter As New OracleDataAdapter(command)
                    connection.Open()
                    adapter.Fill(dataTable)
                End Using
            End Using
        End Using

        Return dataTable
    End Function
    Private Sub SelectedIndexChanged() Handles TabControl1.SelectedIndexChanged
        Dim selectedIndex As Integer = TabControl1.SelectedIndex
        Select Case selectedIndex
            Case 0
                selectDataTab1()
            Case 1
                selectDataTab2()
            Case 2
                selectDataTab3()
            Case 3
                checkFolder()
            Case 4
                selectDataTab5()
        End Select
    End Sub

#Region "tab_1"
    Dim sql_TAB1 = "
    SELECT CASE 
        WHEN LAG(a.URIKEIYAKUNO) OVER (ORDER BY a.URIKEIYAKUNO, a.URIKEIYAKUGYOUNO) = a.URIKEIYAKUNO 
        THEN '〃' 
        ELSE a.URIKEIYAKUNO 
    END AS 売契約No, 
           a.URIKEIYAKUGYOUNO as 行No,c.creationdate as ファイル作成日時,
           a.sunpou as 寸法,
           a.insuu as 員数,
           a.shukkaymd as 出荷日,
           a.W_NO , a.C_NO, 
           a.milsheetno as ミルシートNo,
           a.MAKERCD as メーカーCD,
           b.makerryakumei as メーカー略称,
           b.makermei as メーカー名
    FROM 
          RKM_shipping a, rkm_m_maker b ,FileInfoTable c
   WHERE 
          a.makercd = b.makercd(+) AND a.urikeiyakuno=c.first8chars --and a.milsheetno<>8888888-- a.URIKEIYAKUNO IN ({0}) 
   ORDER BY 
          a.URIKEIYAKUNO, a.URIKEIYAKUGYOUNO"
    Private Sub selectDataTab1()
        'Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        'Dim listKeys As String = ""

        'For Each pdfFile As String In pdfFiles
        '    Dim fileName As String = Path.GetFileNameWithoutExtension(pdfFile)
        '    If fileName.Length >= 8 Then
        '        Dim first8Chars As String = fileName.Substring(0, 8)
        '        ' ファイル名の最初の8文字を取得
        '        If IsNumeric(first8Chars) Then
        '            ' 最初の8文字が数字であるかを確認
        '            listKeys = listKeys + "'" + first8Chars + "',"
        '        End If
        '    End If
        'Next

        'If listKeys.Length > 0 Then
        '    listKeys = listKeys.TrimEnd(","c)
        'End If

        'If Not String.IsNullOrEmpty(listKeys) Then
        '    Dim sqlQuery As String = String.Format(sql_TAB1, listKeys)
        '    Dim resultTable As DataTable = GetData(sqlQuery, connectionStringForYoshu)
        '    DataGridView1.DataSource = Nothing
        '    DataGridView1.DataSource = resultTable
        'End If
        SavePdfFileInfoToOracle()


        Dim resultTable As DataTable = GetData(sql_TAB1, connectionStringForYoshu)
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = resultTable
    End Sub


    Sub SavePdfFileInfoToOracle()
        Dim pdfFiles As String() = Directory.GetFiles(folderPath, "*.pdf")
        Dim listKeys As String = ""

        Using connection As New OracleConnection(connectionStringForYoshu)
            connection.Open()
            Dim query As String = "DELETE FROM FileInfoTable"
            Try
                Using command As New OracleCommand(query, connection)
                    command.ExecuteNonQuery()
                End Using
            Catch ex As OracleException
                ' エラー処理（例: ログへの出力）
                Console.WriteLine($"Oracleエラー: {ex.Message}")
            Catch ex As Exception
                ' その他のエラー処理
                Console.WriteLine($"一般エラー: {ex.Message}")
            End Try

            For Each pdfFile As String In pdfFiles
                Dim fileName As String = Path.GetFileNameWithoutExtension(pdfFile)
                If fileName.Length >= 8 Then
                    Dim first8Chars As String = fileName.Substring(0, 8)
                    ' ファイル名の最初の8文字を取得
                    If IsNumeric(first8Chars) Then
                        ' 最初の8文字が数字であるかを確認
                        listKeys &= "'" & first8Chars & "',"

                        ' ファイルの作成日時を取得
                        Dim creationTime As DateTime = File.GetCreationTime(pdfFile)

                        ' ファイル名と作成日時をデータベースに保存
                        query = "INSERT INTO FileInfoTable (FileName, First8Chars, CreationDate) VALUES (:FileName, :First8Chars, :CreationDate)"
                        Try
                            Using command As New OracleCommand(query, connection)
                                command.Parameters.Add(New OracleParameter(":FileName", fileName))
                                command.Parameters.Add(New OracleParameter(":First8Chars", first8Chars))
                                command.Parameters.Add(New OracleParameter(":CreationDate", creationTime))
                                command.ExecuteNonQuery()
                            End Using
                        Catch ex As OracleException
                            ' エラー処理（例: ログへの出力）
                            Console.WriteLine($"Oracleエラー: {ex.Message}")
                        Catch ex As Exception
                            ' その他のエラー処理
                            Console.WriteLine($"一般エラー: {ex.Message}")
                        End Try
                    End If
                End If
            Next
        End Using
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        selectDataTab1()
    End Sub
#End Region
#Region "tab_2"
    Dim sql_TAB2 = "
    SELECT
    distinct d.KAISHAKBN as 会社区分,
　　d.shimeidisp as 締切　,
　　decode(lengthb(a.tourokutime),7,0||substr( a.tourokutime,1,1) || ':' || substr(a.tourokutime, 2, 2),substr(a.tourokutime,1,2) || ':' || substr(a.tourokutime, 3, 2) ) as 時間,
　　a.USER_YOBI1 as orderNo,
　　a.sunpou as 寸法,
　　a.inzuu as 員数　,
　　a.MAKERCD as メーカーCD　,
　　b.MAKERRYAKUMEI as メーカー略称　,
　　a.c_no　,
　　a.MILSHEETNO as ミルシートNo　,
　　decode(c.MILSHEET_YOUHIKBN,1,'要','不要') as 区分,
　　c.SHUKKAYMD as 出荷日,
　　c.SHUKKAYMD2 as 出荷日２,
　　c.TOUROKUYMD as 出荷データ登録日,
　　c.KOUSHINKAISHAKBN as 更新区分　,
　　c.KOUSHINtantouno as 更新担当

FROM
  RKM_milsheet a,rkm_shipping c,rkm_m_shain d,
   rkm_m_maker b
where a.USER_YOBI1=c.URIKEIYAKUNO and c.KOUSHINTANTOUNO=d.tantouno and
 a.TOUROKUYMD=to_char(sysdate,'yyyymmdd') and a.MAKERCD=b.MAKERCD 
order by
  1,2,3"

    Private Sub selectDataTab2()
        Dim resultTable As DataTable = GetData(sql_TAB2, connectionStringForYoshu)
        DataGridView2.DataSource = Nothing
        DataGridView2.DataSource = resultTable
        Label1.Text = DataGridView2.Rows.Count.ToString() + " 行"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        selectDataTab2()
    End Sub
#End Region
#Region "tab_3"
    Dim sql_TAB3 = "
    SELECT distinct
      s.urikeiyakuno                                        as 豫洲受注NO
    , S.URIKEIYAKUGYOUNO                                    as 豫洲行NO
    , s.shukkaymd                                           as 豫洲出荷日
    , s.sunpou                                              as 豫洲寸法
    , S.INSUU                                               as 豫洲員数
    , A.URIKEIYAKUNO                                        as NS受注NO
    , A.URIKEIYAKUGYOUNO                                    as NS行NO
    , DECODE(
        A.MILSHEETFLG
        , 0
        , '不要'
        , 1
        , '要'
        , 2
        , '同時'
        , 4
        , '切証'
        , 5
        , '切同'
    )                                                       as MS区分
    , decode(c.milsheetno, null, '「未」紐付', '紐付済')    as 紐付
    , decode(a.milsheethakkouflg, 1, '未印刷', 0, '印刷済') as 印刷
    , C.W_NO
    , C.C_NO
FROM
    NS.URIKEIYAKU_M A
    , NS.tr_millprt C
    , NS.rkm_shipping@share4 S
WHERE
    s.MILSHEET_YOUHIKBN = '1'
    AND A.URIKEIYAKUNO = C.URIKEIYAKUNO
    AND A.URIKEIYAKUGYOUNO = C.URIKEIYAKUGYOUNO
    AND replace (a.kyakuchuuno, '-', '') = s.urikeiyakuno
    and s.milsheetno = '9999999'
    and s.USER_SHUKKASAKI like '日鋼ステンレス%'
    and NVL(S.SHUKKAYMD, S.SHUKKAYMD2) > '20211001'
    and a.gaikei = s.gaikei
    and a.nikuatu = s.nikuatu
    and a.NAGASA = s.nagasa
ORDER BY
    s.urikeiyakuno
    , S.URIKEIYAKUGYOUNO
    , S.SUNPOU"

    Private Sub selectDataTab3()
        Dim resultTable As DataTable = GetData(sql_TAB3, connectionStringForNikko)
        DataGridView3.DataSource = Nothing
        DataGridView3.DataSource = resultTable
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        selectDataTab3()
    End Sub
#End Region
#Region "tab_4"
    Public Class FolderInfo
        Public Property LinkFolder As String ' フォルダパス
        Public Property DisplayText As String ' 表示テキスト
        Public Property FileFormat As String ' ファイル形式

        ' コンストラクタ
        Public Sub New(link As String, display As String, format As String)
            LinkFolder = link
            DisplayText = display
            FileFormat = format
        End Sub
    End Class
    Private Sub checkFolder()
        Try
            lbBatch.Text = ""
            Dim folderInfos As New List(Of FolderInfo)()
            ' .iniファイルの読み込み
            If Not IO.File.Exists(fileINI) Then Return

            Dim lines As String() = File.ReadAllLines(fileINI)
            Dim isFolderCheckSection As Boolean = False
            Dim currentLink As String = ""
            Dim currentDisplayText As String = ""
            Dim currentFileFormat As String = ""

            For Each line As String In lines
                line = line.Trim()
                ' [FolderCheckForTab4] セクションの開始を確認
                If line.StartsWith("[FolderCheckForTab4]") Then
                    isFolderCheckSection = True
                    Continue For
                    ' セクションの終了を確認
                ElseIf line.StartsWith("[") AndAlso isFolderCheckSection Then
                    Exit For ' セクション終了後にループを抜ける
                End If

                ' セクション内の行を処理
                If isFolderCheckSection Then
                    If String.IsNullOrWhiteSpace(line) Then
                        Continue For ' 空白行をスキップ
                    End If

                    ' フォルダパスを保存
                    If String.IsNullOrEmpty(currentLink) Then
                        currentLink = line
                    ElseIf String.IsNullOrEmpty(currentDisplayText) Then
                        currentDisplayText = line ' 保存する表示テキスト
                    Else
                        currentFileFormat = line ' 保存するファイルフォーマット

                        ' FolderInfoオブジェクトをリストに追加
                        folderInfos.Add(New FolderInfo(currentLink, currentDisplayText, currentFileFormat))

                        ' リセット
                        currentLink = ""
                        currentDisplayText = ""
                        currentFileFormat = ""
                    End If
                End If
            Next

            ' フォルダのチェックとファイルのカウント
            For Each folderInfo In folderInfos
                ' フォルダが存在するか確認
                If Directory.Exists(folderInfo.LinkFolder) Then
                    Dim count As Integer = 0
                    Dim representativeFiles As New List(Of String)()

                    ' 分割して複数のファイルフォーマットを取得
                    Dim formats As String() = folderInfo.FileFormat.Split(","c)
                    For Each format As String In formats
                        format = format.Trim() ' 不要な空白を削除
                        ' ファイルのカウント
                        Dim files = Directory.GetFiles(folderInfo.LinkFolder, format)
                        Dim visibleFiles = files.Where(Function(f) (File.GetAttributes(f) And FileAttributes.Hidden) = 0).ToArray()
                        count += visibleFiles.Length

                        ' 代表ファイルを取得（先頭1～2個）
                        representativeFiles.AddRange(visibleFiles.Take(1))
                    Next

                    ' フォルダもカウントする場合 (FileFormat = "*.*")
                    If folderInfo.FileFormat.Contains("*.*") Then
                        Dim directories = Directory.GetDirectories(folderInfo.LinkFolder)
                        Dim visibleDirectories = directories.Where(Function(d) (File.GetAttributes(d) And FileAttributes.Hidden) = 0).ToArray()
                        count += visibleDirectories.Length
                    End If

                    If count > 0 Then
                        ' 結果を表示
                        lbBatch.Text += $"{folderInfo.DisplayText}:" + vbNewLine
                        lbBatch.Text += $" {count} ファイル/フォルダ" + vbNewLine

                        If folderInfo.DisplayText.Substring(0, 1) = "★" Then
                            ' 代表ファイル名を追加表示
                            If representativeFiles.Any() Then
                                lbBatch.Text += $"  ⇒ 代表ファイル: {String.Join(", ", representativeFiles.Select(Function(f) Path.GetFileName(f)))}" + vbNewLine
                            End If
                        End If

                        lbBatch.Text += vbNewLine
                    End If
                Else
                    ' フォルダが見つからない場合のエラーメッセージ
                    lbBatch.Text += $"{folderInfo.DisplayText}: " + vbNewLine + $" フォルダが見つかりませんでした。" + vbNewLine + vbNewLine
                    MessageBox.Show($"{folderInfo.DisplayText}: フォルダが見つかりませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Next

        Catch ex As Exception
            ' エラー処理
            MessageBox.Show("エラー: " & ex.Message)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        checkFolder()
    End Sub

    Private Sub CheckMailCount()
        Dim client = New Pop3Client()

        Try
            ' Connect to the POP3 server
            client.Connect(HOST_NAME, PORT, False) ' Use True if the server requires SSL

            ' Authenticate using your email credentials
            client.Authenticate(MAIL_ADDRESS, MAIL_PASSWORD)

            ' Get the number of messages on the server
            Dim messageCount As Integer = client.GetMessageCount()
            Dim message As String = ""

            If messageCount = 0 Then
                message = "ミル子　新着メールはありませんでした"
            Else
                message = $"■ミル子 新着メール {messageCount} 件（山本・上野）{Environment.NewLine}"

                ' Loop through messages and display sender information
                For i As Integer = 1 To Math.Min(5, messageCount) ' Limit to the first 5 messages
                    Dim emailMessage = client.GetMessage(i) ' Get the message
                    Dim sender = emailMessage.Headers.From.Address ' Extract sender's email address
                    Dim senderName = emailMessage.Headers.From.DisplayName ' Extract sender's name
                    Dim subject = emailMessage.Headers.Subject ' Extract email subject

                    ' Add to the message string
                    message &= $"[{i}] {senderName} ({sender}) - 件名: {subject}{Environment.NewLine}"
                Next
            End If

            ' Update the label's text to display the message
            ミル子Label.Text = message

        Catch ex As Exception
            ' Handle any exceptions (e.g., connection issues or authentication failure)
            ミル子Label.Text = $"ミル子　エラー: {ex.Message}"

        Finally
            ' Disconnect from the server to release resources
            client.Disconnect()
        End Try
    End Sub

#End Region
#Region "tab_5"
    Dim sql_TAB5 = ""
    Private Sub selectDataTab5()
        sql_TAB5 = "
        SELECT a.PR_ID,a.MENB   FROM DSP_TBL A  WHERE PR_ID like 'YS02_' and pr_NO > 9 order by 1, pr_no"
        Dim resultTable As DataTable = GetData(sql_TAB5, connectionStringForNikko)
        DataGridView5_1.DataSource = Nothing
        DataGridView5_1.DataSource = resultTable

        sql_TAB5 = "select 'ダブり設定',a.TOKUISAKICD,a.ATUKAITENCD ,a.TODOKESAKICD,null,count(*)
            from m_uritanka_level a where tokuisakicd<>'99999' group by a.TOKUISAKICD,a.ATUKAITENCD ,a.TODOKESAKICD  having count(*)>1 order by 1,2,4"
        resultTable = GetData(sql_TAB5, connectionStringForNikko)
        DataGridView5_2.DataSource = Nothing
        DataGridView5_2.DataSource = resultTable

        sql_TAB5 = "
        SELECT shouhincd, gaikei, nikuatu, nagasa, sum(A.yJ_TOUZAN_I + A.yI_TOUZAN_I) from zfi a group by shouhincd, gaikei, nikuatu, nagasa having sum(A.yJ_TOUZAN_I + A.yI_TOUZAN_I) > 0 minus SELECT shouhincd, gaikei, nikuatu, nagasa, SUM(I_ZAIKO_1 + J_ZAIKO_1 + I_ZAIKO_2 + J_ZAIKO_2 + I_ZAIKO_3 + J_ZAIKO_3 + I_ZAIKO_4 + J_ZAIKO_4 + I_ZAIKO_5 + J_ZAIKO_5 + I_ZAIKO_6 + J_ZAIKO_6 + I_ZAIKO_7 + J_ZAIKO_7 + I_ZAIKO_8 + J_ZAIKO_8 + I_ZAIKO_9 + J_ZAIKO_9 + I_ZAIKO_10 + J_ZAIKO_10 + I_ZAIKO_11 + J_ZAIKO_11 + I_ZAIKO_12 + J_ZAIKO_12 + I_ZAIKO_13 + J_ZAIKO_13 + I_ZAIKO_14 + J_ZAIKO_14 + I_ZAIKO_15 + J_ZAIKO_15 + I_ZAIKO_16 + J_ZAIKO_16 + I_ZAIKO_17 + J_ZAIKO_17 + I_ZAIKO_18 + J_ZAIKO_18 + I_ZAIKO_19 + J_ZAIKO_19 + I_ZAIKO_20 + J_ZAIKO_20 + I_ZAIKO_21 + J_ZAIKO_21 + I_ZAIKO_22 + J_ZAIKO_22 + I_ZAIKO_23 + J_ZAIKO_23 + I_ZAIKO_24 + J_ZAIKO_24 + I_ZAIKO_25 + J_ZAIKO_25 + I_ZAIKO_26 + J_ZAIKO_26 + I_ZAIKO_27 + J_ZAIKO_27 + I_ZAIKO_28 + J_ZAIKO_28 + I_ZAIKO_29 + J_ZAIKO_29 + I_ZAIKO_30 + J_ZAIKO_30 + I_ZAIKO_31 + J_ZAIKO_31 + I_ZAIKO_32 + J_ZAIKO_32 + I_ZAIKO_33 + J_ZAIKO_33 + I_ZAIKO_34 + J_ZAIKO_34 + I_ZAIKO_35 + J_ZAIKO_35 + I_ZAIKO_36 + J_ZAIKO_36 + I_ZAIKO_37 + J_ZAIKO_37 + I_ZAIKO_38 + J_ZAIKO_38 + I_ZAIKO_39 + J_ZAIKO_39 + I_ZAIKO_40 + J_ZAIKO_40) as insuu FROM ZFI_HANBAIKANOU group by shouhincd, gaikei, nikuatu, nagasa having SUM(I_ZAIKO_1 + J_ZAIKO_1 + I_ZAIKO_2 + J_ZAIKO_2 + I_ZAIKO_3 + J_ZAIKO_3 + I_ZAIKO_4 + J_ZAIKO_4 + I_ZAIKO_5 + J_ZAIKO_5 + I_ZAIKO_6 + J_ZAIKO_6 + I_ZAIKO_7 + J_ZAIKO_7 + I_ZAIKO_8 + J_ZAIKO_8 + I_ZAIKO_9 + J_ZAIKO_9 + I_ZAIKO_10 + J_ZAIKO_10 + I_ZAIKO_11 + J_ZAIKO_11 + I_ZAIKO_12 + J_ZAIKO_12 + I_ZAIKO_13 + J_ZAIKO_13 + I_ZAIKO_14 + J_ZAIKO_14 + I_ZAIKO_15 + J_ZAIKO_15 + I_ZAIKO_16 + J_ZAIKO_16 + I_ZAIKO_17 + J_ZAIKO_17 + I_ZAIKO_18 + J_ZAIKO_18 + I_ZAIKO_19 + J_ZAIKO_19 + I_ZAIKO_20 + J_ZAIKO_20 + I_ZAIKO_21 + J_ZAIKO_21 + I_ZAIKO_22 + J_ZAIKO_22 + I_ZAIKO_23 + J_ZAIKO_23 + I_ZAIKO_24 + J_ZAIKO_24 + I_ZAIKO_25 + J_ZAIKO_25 + I_ZAIKO_26 + J_ZAIKO_26 + I_ZAIKO_27 + J_ZAIKO_27 + I_ZAIKO_28 + J_ZAIKO_28 + I_ZAIKO_29 + J_ZAIKO_29 + I_ZAIKO_30 + J_ZAIKO_30 + I_ZAIKO_31 + J_ZAIKO_31 + I_ZAIKO_32 + J_ZAIKO_32 + I_ZAIKO_33 + J_ZAIKO_33 + I_ZAIKO_34 + J_ZAIKO_34 + I_ZAIKO_35 + J_ZAIKO_35 + I_ZAIKO_36 + J_ZAIKO_36 + I_ZAIKO_37 + J_ZAIKO_37 + I_ZAIKO_38 + J_ZAIKO_38 + I_ZAIKO_39 + J_ZAIKO_39 + I_ZAIKO_40 + J_ZAIKO_40) <> 0 union all SELECT shouhincd, gaikei, nikuatu, nagasa, SUM(I_ZAIKO_1 + J_ZAIKO_1 + I_ZAIKO_2 + J_ZAIKO_2 + I_ZAIKO_3 + J_ZAIKO_3 + I_ZAIKO_4 + J_ZAIKO_4 + I_ZAIKO_5 + J_ZAIKO_5 + I_ZAIKO_6 + J_ZAIKO_6 + I_ZAIKO_7 + J_ZAIKO_7 + I_ZAIKO_8 + J_ZAIKO_8 + I_ZAIKO_9 + J_ZAIKO_9 + I_ZAIKO_10 + J_ZAIKO_10 + I_ZAIKO_11 + J_ZAIKO_11 + I_ZAIKO_12 + J_ZAIKO_12 + I_ZAIKO_13 + J_ZAIKO_13 + I_ZAIKO_14 + J_ZAIKO_14 + I_ZAIKO_15 + J_ZAIKO_15 + I_ZAIKO_16 + J_ZAIKO_16 + I_ZAIKO_17 + J_ZAIKO_17 + I_ZAIKO_18 + J_ZAIKO_18 + I_ZAIKO_19 + J_ZAIKO_19 + I_ZAIKO_20 + J_ZAIKO_20 + I_ZAIKO_21 + J_ZAIKO_21 + I_ZAIKO_22 + J_ZAIKO_22 + I_ZAIKO_23 + J_ZAIKO_23 + I_ZAIKO_24 + J_ZAIKO_24 + I_ZAIKO_25 + J_ZAIKO_25 + I_ZAIKO_26 + J_ZAIKO_26 + I_ZAIKO_27 + J_ZAIKO_27 + I_ZAIKO_28 + J_ZAIKO_28 + I_ZAIKO_29 + J_ZAIKO_29 + I_ZAIKO_30 + J_ZAIKO_30 + I_ZAIKO_31 + J_ZAIKO_31 + I_ZAIKO_32 + J_ZAIKO_32 + I_ZAIKO_33 + J_ZAIKO_33 + I_ZAIKO_34 + J_ZAIKO_34 + I_ZAIKO_35 + J_ZAIKO_35 + I_ZAIKO_36 + J_ZAIKO_36 + I_ZAIKO_37 + J_ZAIKO_37 + I_ZAIKO_38 + J_ZAIKO_38 + I_ZAIKO_39 + J_ZAIKO_39 + I_ZAIKO_40 + J_ZAIKO_40) as insuu FROM ZFI_HANBAIKANOU group by shouhincd, gaikei, nikuatu, nagasa having SUM(I_ZAIKO_1 + J_ZAIKO_1 + I_ZAIKO_2 + J_ZAIKO_2 + I_ZAIKO_3 + J_ZAIKO_3 + I_ZAIKO_4 + J_ZAIKO_4 + I_ZAIKO_5 + J_ZAIKO_5 + I_ZAIKO_6 + J_ZAIKO_6 + I_ZAIKO_7 + J_ZAIKO_7 + I_ZAIKO_8 + J_ZAIKO_8 + I_ZAIKO_9 + J_ZAIKO_9 + I_ZAIKO_10 + J_ZAIKO_10 + I_ZAIKO_11 + J_ZAIKO_11 + I_ZAIKO_12 + J_ZAIKO_12 + I_ZAIKO_13 + J_ZAIKO_13 + I_ZAIKO_14 + J_ZAIKO_14 + I_ZAIKO_15 + J_ZAIKO_15 + I_ZAIKO_16 + J_ZAIKO_16 + I_ZAIKO_17 + J_ZAIKO_17 + I_ZAIKO_18 + J_ZAIKO_18 + I_ZAIKO_19 + J_ZAIKO_19 + I_ZAIKO_20 + J_ZAIKO_20 + I_ZAIKO_21 + J_ZAIKO_21 + I_ZAIKO_22 + J_ZAIKO_22 + I_ZAIKO_23 + J_ZAIKO_23 + I_ZAIKO_24 + J_ZAIKO_24 + I_ZAIKO_25 + J_ZAIKO_25 + I_ZAIKO_26 + J_ZAIKO_26 + I_ZAIKO_27 + J_ZAIKO_27 + I_ZAIKO_28 + J_ZAIKO_28 + I_ZAIKO_29 + J_ZAIKO_29 + I_ZAIKO_30 + J_ZAIKO_30 + I_ZAIKO_31 + J_ZAIKO_31 + I_ZAIKO_32 + J_ZAIKO_32 + I_ZAIKO_33 + J_ZAIKO_33 + I_ZAIKO_34 + J_ZAIKO_34 + I_ZAIKO_35 + J_ZAIKO_35 + I_ZAIKO_36 + J_ZAIKO_36 + I_ZAIKO_37 + J_ZAIKO_37 + I_ZAIKO_38 + J_ZAIKO_38 + I_ZAIKO_39 + J_ZAIKO_39 + I_ZAIKO_40 + J_ZAIKO_40) <> 0 minus SELECT shouhincd, gaikei, nikuatu, nagasa, sum(A.yJ_TOUZAN_I + A.yI_TOUZAN_I) from zfi a group by shouhincd, gaikei, nikuatu, nagasa having sum(A.yJ_TOUZAN_I + A.yI_TOUZAN_I) > 0"
        resultTable = GetData(sql_TAB5, connectionStringForNikko)
        DataGridView5_3.DataSource = Nothing
        DataGridView5_3.DataSource = resultTable

        CheckMailCount()
        ' Label1.Text = DataGridView2.Rows.Count.ToString() + " 行"
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        selectDataTab5()
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        PatternKensaku起動("00005,00006,00007,00008,00009,00010")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        PatternKensaku起動(TextBox実行Key.Text)
    End Sub

    Private Shared Sub PatternKensaku起動(key As String)
        ' プロセスの情報を設定
        Dim startInfo As New ProcessStartInfo()
        startInfo.FileName = "C:\\新システム\\パターン検索\\test\\Global_Batch"
        startInfo.Arguments = key
        startInfo.UseShellExecute = False ' コンソール出力を取得しない場合はTrueに変更可能

        Try
            ' プロセスを開始
            Dim process As Process = Process.Start(startInfo)

            ' プロセスが終了するのを待つ場合
            ' process.WaitForExit()
        Catch ex As Exception
            ' エラーメッセージを表示
            MessageBox.Show("エラーが発生しました: " & ex.Message)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Process.Start("C:\新システム\棚卸関係\棚卸出欠反映処理\棚卸出欠反映")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Process.Start("https://nikko-sus.cybozu.com/o/ag.cgi?page=BulletinView&bid=35420&gid=0&cid=9754&cp=blc&wt=")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Process.Start("C:\rakumiru\5001\IMAGE")
        Process.Start("Q:\SHARE\5001\IMAGE")
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        SavePdfFileInfoToOracle()
    End Sub
End Class


#End Region

