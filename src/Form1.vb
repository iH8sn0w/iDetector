Imports System
Imports System.Management
Imports System.Windows.Forms
Public Class Form1
    Public iBoot As String = ""
    Public CPID As String = ""
    Sub Delay(ByVal dblSecs As Double)

        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop

    End Sub
    Function HighlightWords(ByVal rtb As RichTextBox, ByVal sFindString As String, ByVal lColor As System.Drawing.Color) As Integer

        Dim iFoundPos As Integer 'Position of first character of match
        Dim iFindLength As Integer       'Length of string to find
        Dim iOriginalSelStart As Integer
        Dim iOriginalSelLength As Integer
        Dim iMatchCount As Integer      'Number of matches

        'Save the insertion points current location and length
        iOriginalSelStart = rtb.SelectionStart
        iOriginalSelLength = rtb.SelectionLength

        'Cache the length of the string to find
        iFindLength = Len(sFindString) + 16

        'Attempt to find the first match
        iFoundPos = rtb.Find(sFindString, 0, RichTextBoxFinds.NoHighlight)
        While iFoundPos > 0
            iMatchCount = iMatchCount + 1

            console.SelectionStart = iFoundPos
            'The SelLength property is set to 0 as soon as you change SelStart
            console.SelectionLength = iFindLength
            'rtb.SelectionBackColor = lColor

            console.Select(iFoundPos, iFindLength - 8)
            iBoot = console.SelectedText
            'MsgBox(iBoot)
            'Attempt to find the next match
            iFoundPos = rtb.Find(sFindString, iFoundPos + iFindLength, RichTextBoxFinds.NoHighlight)
        End While

        'Restore the insertion point to its original location and length
        rtb.SelectionStart = iOriginalSelStart
        rtb.SelectionLength = iOriginalSelLength

        'Return the number of matches
        HighlightWords = iMatchCount
    End Function
    Function HighlightWords2(ByVal rtb As RichTextBox, ByVal sFindString2 As String, ByVal lColor2 As System.Drawing.Color) As Integer

        Dim iFoundPos2 As Integer 'Position of first character of match
        Dim iFindLength2 As Integer       'Length of string to find
        Dim iOriginalSelStart2 As Integer
        Dim iOriginalSelLength2 As Integer
        Dim iMatchCount2 As Integer      'Number of matches

        'Save the insertion points current location and length
        iOriginalSelStart2 = rtb.SelectionStart
        iOriginalSelLength2 = rtb.SelectionLength

        'Cache the length of the string to find
        iFindLength2 = Len(sFindString2)

        'Attempt to find the first match
        iFoundPos2 = rtb.Find(sFindString2, 0, RichTextBoxFinds.NoHighlight)
        While iFoundPos2 > 0
            iMatchCount2 = iMatchCount2 + 1

            console.SelectionStart = iFoundPos2
            'The SelLength property is set to 0 as soon as you change SelStart
            console.SelectionLength = iFindLength2
            'rtb.SelectionBackColor = lColor2

            console.Select(iFoundPos2 + 5, iFindLength2 - 1)
            CPID = console.SelectedText
            'MsgBox(CPID)
            'Attempt to find the next match
            iFoundPos2 = rtb.Find(sFindString2, iFoundPos2 + iFindLength2, RichTextBoxFinds.NoHighlight)
        End While

        'Restore the insertion point to its original location and length
        rtb.SelectionStart = iOriginalSelStart2
        rtb.SelectionLength = iOriginalSelLength2

        'Return the number of matches
        HighlightWords2 = iMatchCount2
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button1.Enabled = False
        Button1.Text = "Searching for DFU..."
        Dim searcher As New ManagementObjectSearcher( _
                    "root\CIMV2", _
                    "SELECT * FROM Win32_PnPDevice")

        For Each queryObj As ManagementObject In searcher.Get()
            console.Text += ("SystemElement: {0}" & queryObj("SystemElement"))
        Next
        If console.Text.Contains("ECID") Then
            Dim icountMatch2 As Integer
            icountMatch2 = HighlightWords2(console, "CPID:", System.Drawing.Color.Red)
            If CPID = "8920" Then
                Call GoGoGadgetiBoot()
            Else
                MsgBox("Another device other than a 3GS was detected! Please connect ONLY a 3GS.", MsgBoxStyle.Critical)
            End If
        Else
            MsgBox("No Device was Detected!")
            Button1.Enabled = True
            Button1.Text = "Is my Bootrom Old or New?"
        End If
    End Sub
    Public Sub GoGoGadgetiBoot()
        Dim icountMatch As Integer
        icountMatch = HighlightWords(console, "IBOOT", System.Drawing.Color.Red)
        If iBoot = "IBOOT-359.3]""" Then
            MsgBox("Your iPhone 3G[S] contains the Old Bootrom." & Chr(13) & Chr(13) & "You can Exit DFU Mode by holding the Power + Home button continuously.")
            Button1.Enabled = True
            Button1.Text = "Is my Bootrom Old or New?"
        Else
            MsgBox("Your iPhone 3G[S] contains the NEW Bootrom." & Chr(13) & Chr(13) & "You can Exit DFU Mode by holding the Power + Home button continuously.")
            Button1.Enabled = True
            Button1.Text = "Is my Bootrom Old or New?"
        End If
    End Sub
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Process.Start("http://www.youtube.com/watch?v=bITIiGswjFI")
    End Sub
End Class
