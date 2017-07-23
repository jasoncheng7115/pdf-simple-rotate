Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser

Public Class Form1

    Dim DragDropFileName As String = ""
    Dim NewTargetPdfFileName As String = ""
    Dim SelectFileName As String = ""
    Dim SelectFileNameText As String = ""

    ''' <summary>
    ''' 處理檔案拖曳行為
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop

        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

        If files.Count > 1 Then

            '超過數量，不予處理
            LinkLabel1.Visible = False
            MsgBox("一次只能轉換一個檔案", MsgBoxStyle.Exclamation, "提醒")

        Else

            '記錄與顯示相關檔名資訊
            LinkLabel1.Visible = False
            DragDropFileName = files(0).ToString()
            SelectFileName = files(0).ToString()
            SelectFileNameText = "您已選擇 '" & files(0).ToString() & "' 檔案進行轉向"
            TextBox1.Text = SelectFileNameText

        End If


    End Sub

    ''' <summary>
    ''' 起用拖曳進入行為
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop, False) Then
            e.Effect = DragDropEffects.All
        End If

    End Sub

    ''' <summary>
    ''' 表單載入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LinkLabel1.Visible = False
        LinkLabel1.Text = "[開啟檔案]"

    End Sub

    ''' <summary>
    ''' 進行轉向處理
    ''' </summary>
    ''' <param name="SourcePdfFileName">來源檔名與路徑</param>
    ''' <param name="rotation">轉向角度</param>
    ''' <remarks></remarks>
    Sub DoRotate(SourcePdfFileName As String, rotation As Integer)

        '沒有檔案
        If DragDropFileName = "" Then
            MsgBox("您尚未拖曳要轉向之檔案進來", MsgBoxStyle.Exclamation, "提醒")
            Exit Sub
        End If

        Try

            '組成新檔名
            NewTargetPdfFileName = ""
            NewTargetPdfFileName &= IO.Path.GetDirectoryName(SourcePdfFileName) & "\" & IO.Path.GetFileNameWithoutExtension(SourcePdfFileName) & "_" & rotation.ToString & IO.Path.GetExtension(SourcePdfFileName)

            '讀入來源檔
            Dim reader As New PdfReader(SourcePdfFileName)
            Dim pagesCount As Integer = reader.NumberOfPages()

            For n As Integer = 1 To pagesCount

                TextBox1.Text = SelectFileNameText & vbCrLf & vbCrLf & "轉換進度： " & n & "/" & pagesCount & " 頁"

                Dim page As PdfDictionary = reader.GetPageN(n)

                '轉向並回寫至頁面
                Dim rotate As PdfNumber = page.GetAsNumber(PdfName.ROTATE)
                page.Put(PdfName.ROTATE, New PdfNumber(rotation))

                '釋放一下 CPU
                Threading.Thread.Sleep(1)

            Next

            '寫入新檔
            Dim Stamper As PdfStamper = New PdfStamper(reader, New IO.FileStream(NewTargetPdfFileName, IO.FileMode.Create))

            Stamper.Close()
            reader.Close()

            TextBox1.Text = SelectFileNameText & vbCrLf & vbCrLf & "轉換完成，檔案 '" & NewTargetPdfFileName & "'"
            LinkLabel1.Visible = True

        Catch ex As Exception

            '顯示錯誤結果
            MsgBox("寫入 '" & NewTargetPdfFileName & "' 時發生錯誤：" & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "錯誤")
            TextBox1.Text = SelectFileNameText & vbCrLf & vbCrLf & "轉換發生問題，請注意目的地是否無權限或被開啟中，或是非 PDF 格式。"

        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        CType(sender, Button).Enabled = False
        DoRotate(DragDropFileName, 90)
        CType(sender, Button).Enabled = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        CType(sender, Button).Enabled = False
        DoRotate(DragDropFileName, 270)
        CType(sender, Button).Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        CType(sender, Button).Enabled = False
        DoRotate(DragDropFileName, 180)
        CType(sender, Button).Enabled = True

    End Sub

    ''' <summary>
    ''' 開啟轉換完成之檔案
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        LinkLabel1.LinkVisited = True
        System.Diagnostics.Process.Start(NewTargetPdfFileName)
    End Sub

End Class
