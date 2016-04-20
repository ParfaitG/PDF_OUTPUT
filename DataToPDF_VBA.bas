
''''''''''''''''''
'' MS ACCESS
''''''''''''''''''
Option Compare Database
Option Explicit

Private Sub CSVImport_Click()
On Error GoTo ErrHandle
    Dim strFile As String
    
    strFile = Application.CurrentProject.Path
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM EnergyData"
    DoCmd.TransferText acImportDelim, , "EnergyData", strFile & "\DATA\EnergyConsumptionBySector1949-2015.csv", True
    DoCmd.SetWarnings True
    
    MsgBox "Successfully imported CSV data!", vbInformation
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Exit Sub
End Sub


Private Sub PDFExport_Click()
On Error GoTo ErrHandle
    Dim rawDoc As Object, xslDoc As Object, newDoc As Object
    Dim xmlstr As String, xslstr As String, todayDate As String
    Dim strResult As String
    Dim fso As Object, ofile As Object
    
    Dim execpath As String, retVal As Integer
    Dim execstyle As Integer: execstyle = 1
    Dim waitTillComplete As Boolean: waitTillComplete = True
    Dim shell As Object
    
    xmlstr = Application.CurrentProject.Path & "\DATA\Output_ACC.xml"
        
    ' EXPORT TABLE TO XML FORMAT
    Application.ExportXML acExportQuery, "EnergyPivot", xmlstr
        
    ' TRANFORM XML TO HTML
    Set rawDoc = CreateObject("MSXML2.DOMDocument")
    Set xslDoc = CreateObject("MSXML2.DOMDocument")
    Set newDoc = CreateObject("MSXML2.DOMDocument")

    rawDoc.async = False
    rawDoc.Load Application.CurrentProject.Path & "\DATA\Output_ACC.xml"

    xslDoc.async = False
    xslDoc.loadXML DLookup("Script", "ScriptsData", "Purpose='XML to HTML Transform'")
    strResult = rawDoc.transformNode(xslDoc)
    
    ' OUTPUT TRANSFORMATION TO FILE
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ofile = fso.CreateTextFile(Application.CurrentProject.Path & "\DATA\Output_ACC.html")
    ofile.WriteLine strResult
    ofile.Close
        
    ' CONVERT TO PDF
    Set shell = VBA.CreateObject("WScript.Shell")
    execpath = "wkhtmltopdf.exe -O landscape """ & Application.CurrentProject.Path & "\DATA\Output_ACC.html""" _
                  & " """ & Application.CurrentProject.Path & "\DATA\Output_ACC.pdf"""
    retVal = shell.Run(execpath, execstyle, waitTillComplete)
    
    MsgBox "Successfully processed CSV data to PDF!", vbInformation
    
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Exit Sub
    
End Sub


''''''''''''''''''
'' MS EXCEL
''''''''''''''''''
Option Explicit

Public Sub RunAppRoutines()
    Call DataHandle
    Call TransposeData
    Call htmlExport
End Sub

Public Sub DataHandle()
On Error GoTo ErrHandle
    Dim qt As QueryTable
    Dim wkb As Workbook, dwks As Worksheet, twks As Worksheet
    Dim strPath As String
    Dim i As Long, j As Long, yr As Integer, mo As Integer
    
    Application.ScreenUpdating = False
    
    strPath = Application.ActiveWorkbook.path
    Set dwks = ThisWorkbook.Worksheets("DATA")
    
    dwks.Columns("A:I").EntireColumn.Delete xlShiftToLeft
    
    ' READ CSV
    With dwks.QueryTables.Add(Connection:="TEXT;" & strPath & "\DATA\EnergyConsumptionBySector1949-2015.csv", _
        Destination:=dwks.Cells(1, 1))
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False

            .Refresh BackgroundQuery:=False
    End With
    
    For Each qt In dwks.QueryTables
        qt.Delete
    Next qt
    
    ' YEAR AND MONTH
    dwks.Columns("C:D").EntireColumn.Insert
    dwks.Range("C1") = "Year"
    dwks.Range("D1") = "Month"
    
    For i = 2 To dwks.UsedRange.Rows.Count
        dwks.Range("C" & i) = Left(dwks.Range("B" & i), 4)
        dwks.Range("D" & i) = Mid(dwks.Range("B" & i), 5, 2)
    Next i
    
    Application.ScreenUpdating = True
    Set dwks = Nothing
    Set wkb = Nothing
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "RUN-TIME ERROR"
    Exit Sub
        
End Sub

Public Sub TransposeData()
On Error GoTo ErrHandle
    Dim twks As Worksheet, dwks As Worksheet
    Dim i As Long, j As Long, yr As Integer, mo As Integer

    Application.ScreenUpdating = False
    Set dwks = ThisWorkbook.Worksheets("DATA")
    
    Set twks = ThisWorkbook.Worksheets("TRANSPOSED")
    twks.Columns("A:N").EntireColumn.Delete xlShiftToLeft
    
    twks.Range("A" & 1) = "Year"
    twks.Range("B" & 1) = "Month"
    twks.Range("C" & 1) = "Residential Sector Primary"
    twks.Range("D" & 1) = "Residential Sector Total"
    twks.Range("E" & 1) = "Commercial Sector Primary"
    twks.Range("F" & 1) = "Commercial Sector Total"
    twks.Range("G" & 1) = "Industrial Sector Primary"
    twks.Range("H" & 1) = "Industrial Sector Total"
    twks.Range("I" & 1) = "Transportation Sector Primary"
    twks.Range("J" & 1) = "Transportation Sector Total"
    twks.Range("K" & 1) = "Electric Power Sector Primary"
    twks.Range("L" & 1) = "Energy Consumption Balancing Item"
    twks.Range("M" & 1) = "Grant Total Primary Consumption"
    
    j = 2
    For yr = 1949 To 2015
        For mo = 1 To 13
        
            For i = 2 To dwks.UsedRange.Rows.Count
                If dwks.Range("C" & i) = yr And dwks.Range("D" & i) = mo Then
                
                    twks.Range("A" & j) = yr
                    If mo = 13 Then twks.Range("B" & j) = "Total" Else twks.Range("B" & j) = MonthName(mo)
            
                    Select Case dwks.Range("G" & i)
                    
                        Case "Primary Energy Consumed by the Residential Sector": twks.Range("C" & j) = dwks.Range("E" & i)
                        Case "Total Energy Consumed by the Residential Sector": twks.Range("D" & j) = dwks.Range("E" & i)
                        Case "Primary Energy Consumed by the Commercial Sector": twks.Range("E" & j) = dwks.Range("E" & i)
                        Case "Total Energy Consumed by the Commercial Sector": twks.Range("F" & j) = dwks.Range("E" & i)
                        Case "Primary Energy Consumed by the Industrial Sector": twks.Range("G" & j) = dwks.Range("E" & i)
                        Case "Total Energy Consumed by the Industrial Sector": twks.Range("H" & j) = dwks.Range("E" & i)
                        Case "Primary Energy Consumed by the Transportation Sector": twks.Range("I" & j) = dwks.Range("E" & i)
                        Case "Total Energy Consumed by the Transportation Sector": twks.Range("J" & j) = dwks.Range("E" & i)
                        Case "Primary Energy Consumed by the Electric Power Sector": twks.Range("K" & j) = dwks.Range("E" & i)
                        Case "Energy Consumption Balancing Item": twks.Range("L" & j) = dwks.Range("E" & i)
                        Case "Primary Energy Consumption Total": twks.Range("M" & j) = dwks.Range("E" & i)
                        
                    End Select
                    
                
                End If
                
            Next i
            If twks.Range("A" & j) <> "" Then j = j + 1
        Next mo
        
    Next yr
    
    Set dwks = Nothing
    Set twks = Nothing
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "RUN-TIME ERROR"
    Exit Sub
    
End Sub

Sub htmlExport()
On Error GoTo ErrHandle
    Dim doc As New MSXML2.DOMDocument60, xslDoc As New MSXML2.DOMDocument60, newDoc As New MSXML2.DOMDocument60
    Dim execpath As String, retVal As Integer
    Dim execstyle As Integer: execstyle = 1
    Dim waitTillComplete As Boolean: waitTillComplete = True
    Dim shell As Object
        
    ' ELEMENT OBJECTS
    Dim root As IXMLDOMElement, headNode As IXMLDOMElement, bodynode As IXMLDOMElement
    Dim styleNode As IXMLDOMElement, imgNode As IXMLDOMElement, h1Node As IXMLDOMElement
    Dim yearTitleNode As IXMLDOMElement, yrimgNode As IXMLDOMElement
    Dim tableNode As IXMLDOMElement, trNode As IXMLDOMElement, tdNode As IXMLDOMElement
    Dim divNode As IXMLDOMElement
    Dim strVal As String
    
    ' ATTRIBUTE OBJECTS
    Dim styletypeAttrib As IXMLDOMAttribute, stylemediaAttrib As IXMLDOMAttribute
    Dim h2class As IXMLDOMAttribute, yrclassAttrib As IXMLDOMAttribute
    Dim imgsrcAttrib As IXMLDOMAttribute, imgaltAttrib As IXMLDOMAttribute
    Dim yrimgsrcAttrib As IXMLDOMAttribute, yrimgaltAttrib As IXMLDOMAttribute
    Dim trheadAttrib As IXMLDOMAttribute, trclassevenAttrib As IXMLDOMAttribute
    Dim divclass As IXMLDOMAttribute
    Dim i As Long, j As Long, yr As Long
    
    ' DECLARE XML DOC OBJECT '
    Set root = doc.createElement("html")
    doc.appendChild root
            
    Set headNode = doc.createElement("head")
    root.appendChild headNode
    
    Set styleNode = doc.createElement("style")
        strVal = vbNewLine & "     body{" & vbNewLine
        strVal = strVal & "          margin:15px; padding: 20px;" & vbNewLine
        strVal = strVal & "          font-family:Arial, Helvetica, sans-serif; font-size:88%; " & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     h1, h2 {" & vbNewLine
        strVal = strVal & "          font:Arial black; color: #383838;" & vbNewLine
        strVal = strVal & "          valign: top;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     .yeartitle{" & vbNewLine
        strVal = strVal & "          page-break-before: always;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     img {" & vbNewLine
        strVal = strVal & "          float: right;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     table, tr, td, th, thead, tbody, tfoot {" & vbNewLine
        strVal = strVal & "          page-break-inside: avoid !important;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     table{" & vbNewLine
        strVal = strVal & "          width:100%; font-size:13px;" & vbNewLine
        strVal = strVal & "          border-collapse:collapse;" & vbNewLine
        strVal = strVal & "          text-align: right;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     th{ color: #383838 ; padding:2px; text-align:right; }" & vbNewLine
        strVal = strVal & "     td{ padding: 2px 5px 2px 5px; }" & vbNewLine
        strVal = strVal & "     tr.headerrow{" & vbNewLine
        strVal = strVal & "          border-bottom: 2px solid #A8A8A8;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     tr.even{" & vbNewLine
        strVal = strVal & "          background-color: #F0F0F0;"
        strVal = strVal & "     }" & vbNewLine
        strVal = strVal & "     .footer{" & vbNewLine
        strVal = strVal & "          text-align: right;" & vbNewLine
        strVal = strVal & "          color: #A8A8A8;" & vbNewLine
        strVal = strVal & "          font-size: 12px;" & vbNewLine
        strVal = strVal & "          margin: 10px;" & vbNewLine
        strVal = strVal & "     }" & vbNewLine
        
    styleNode.Text = strVal
    headNode.appendChild styleNode
    
        Set styletypeAttrib = doc.createAttribute("type")
        styletypeAttrib.Value = "text/css"
        styleNode.setAttributeNode styletypeAttrib
        
        Set stylemediaAttrib = doc.createAttribute("media")
        stylemediaAttrib.Value = "all"
        styleNode.setAttributeNode stylemediaAttrib
        
    
    Set bodynode = doc.createElement("body")
    root.appendChild bodynode
                
    Set h1Node = doc.createElement("h1")
    h1Node.Text = "U.S. Energy Consumption 1949 - 2015"
    bodynode.appendChild h1Node
                
    Set imgNode = doc.createElement("img")
    h1Node.appendChild imgNode
                                
        Set imgsrcAttrib = doc.createAttribute("src")
        imgsrcAttrib.Value = "EnergyIcon.png"
        imgNode.setAttributeNode imgsrcAttrib
        
        Set imgaltAttrib = doc.createAttribute("alt")
        imgaltAttrib.Value = "energy icon"
        imgNode.setAttributeNode imgaltAttrib

    ' TABLE NODE '
    Set tableNode = doc.createElement("table")
    bodynode.appendChild tableNode
    
    Set trNode = runHeaders(doc, tableNode)
    tableNode.appendChild trNode
                
    ' DATA ROWS '
    For i = 2 To ThisWorkbook.Worksheets("TRANSPOSED").UsedRange.Rows.Count
        
        If ThisWorkbook.Worksheets("TRANSPOSED").Range("A" & i) <= 1972 Then
        
            Set trNode = doc.createElement("tr")
            tableNode.appendChild trNode
            
            If i Mod 2 <> 0 Then
                Set trclassevenAttrib = doc.createAttribute("class")
                trclassevenAttrib.Value = "even"
                trNode.setAttributeNode trclassevenAttrib
            End If
        
            For j = 1 To 13
                Set tdNode = doc.createElement("td")
                tdNode.Text = ThisWorkbook.Worksheets("TRANSPOSED").Cells(i, j)
                trNode.appendChild tdNode
            Next j
        End If
    Next i
    
    Set divNode = doc.createElement("div")
    divNode.Text = "Source: EIA - U.S. Department of Energy"
    bodynode.appendChild divNode
    
    Set divclass = doc.createAttribute("class")
    divclass.Value = "footer"
    divNode.setAttributeNode divclass
    
    ' DATA TABLES '
    For yr = 1973 To 2015
        
        Set yearTitleNode = doc.createElement("h2")
        yearTitleNode.Text = yr
        bodynode.appendChild yearTitleNode
        
            Set yrclassAttrib = doc.createAttribute("class")
            yrclassAttrib.Value = "yeartitle"
            yearTitleNode.setAttributeNode yrclassAttrib
        
        Set yrimgNode = doc.createElement("img")
        yearTitleNode.appendChild yrimgNode

            Set yrimgsrcAttrib = doc.createAttribute("src")
            yrimgsrcAttrib.Value = "EnergyIcon.png"
            yrimgNode.setAttributeNode yrimgsrcAttrib
            
            Set yrimgaltAttrib = doc.createAttribute("alt")
            yrimgaltAttrib.Value = "energy icon"
            yrimgNode.setAttributeNode yrimgaltAttrib
                                    
        Set tableNode = doc.createElement("table")
        bodynode.appendChild tableNode
    
        Set trNode = runHeaders(doc, tableNode)
        tableNode.appendChild trNode
        
        For i = 2 To ThisWorkbook.Worksheets("TRANSPOSED").UsedRange.Rows.Count
            If ThisWorkbook.Worksheets("TRANSPOSED").Range("A" & i) = yr Then
                            
                Set trNode = doc.createElement("tr")
                tableNode.appendChild trNode
                
                If i Mod 2 = 0 Then
                    Set trclassevenAttrib = doc.createAttribute("class")
                    trclassevenAttrib.Value = "even"
                    trNode.setAttributeNode trclassevenAttrib
                End If
            
                For j = 1 To 13
                    Set tdNode = doc.createElement("td")
                    tdNode.Text = ThisWorkbook.Worksheets("TRANSPOSED").Cells(i, j)
                    trNode.appendChild tdNode
                Next j
                
            End If
        Next i
        Set divNode = doc.createElement("div")
        divNode.Text = "Source: EIA - U.S. Department of Energy"
        bodynode.appendChild divNode
        
        Set divclass = doc.createAttribute("class")
        divclass.Value = "footer"
        divNode.setAttributeNode divclass
    Next yr
    
    ' PRETTY PRINT RAW OUTPUT '
    xslDoc.LoadXML "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>" _
            & "<xsl:stylesheet version=" & Chr(34) & "1.0" & Chr(34) _
            & "                xmlns:xsl=" & Chr(34) & "http://www.w3.org/1999/XSL/Transform" & Chr(34) & ">" _
            & "<xsl:strip-space elements=" & Chr(34) & "*" & Chr(34) & " />" _
            & "<xsl:output method=" & Chr(34) & "xml" & Chr(34) & " indent=" & Chr(34) & "yes" & Chr(34) & "" _
            & "            encoding=" & Chr(34) & "UTF-8" & Chr(34) & "/>" _
            & " <xsl:template match=" & Chr(34) & "node() | @*" & Chr(34) & ">" _
            & "  <xsl:copy>" _
            & "   <xsl:apply-templates select=" & Chr(34) & "node() | @*" & Chr(34) & " />" _
            & "  </xsl:copy>" _
            & " </xsl:template>" _
            & "</xsl:stylesheet>"

    xslDoc.async = False
    doc.transformNodeToObject xslDoc, newDoc
    newDoc.Save ActiveWorkbook.path & "\DATA\Output_XL.html"
            
    ' CONVERT TO PDF
    Set shell = VBA.CreateObject("WScript.Shell")
    execpath = "wkhtmltopdf.exe -O landscape """ & ActiveWorkbook.path & "\DATA\Output_XL.html""" _
                  & " """ & ActiveWorkbook.path & "\DATA\Output_XL.pdf"""
    retVal = shell.Run(execpath, execstyle, waitTillComplete)

    MsgBox "Successfully processed CSV data to PDF!", vbInformation
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "RUN-TIME ERROR"
    Exit Sub
    
End Sub

Public Function runHeaders(docobj As MSXML2.DOMDocument60, tableobj As IXMLDOMElement) As IXMLDOMElement
    Dim trNode As IXMLDOMElement, thNode As IXMLDOMElement
    Dim trClass As IXMLDOMAttribute
    Dim i As Variant
    
    Set trNode = docobj.createElement("tr")
    tableobj.appendChild trNode
    
    Set trClass = docobj.createAttribute("class")
    trClass.Value = "headerrow"
    trNode.setAttributeNode trClass
    
    For Each i In Array("Year", "Month", _
                       "Residential Sector Primary", "Residential Sector Total", _
                       "Commercial Sector Primary", "Commercial Sector Total", _
                       "Industrial Sector Primary", "Industrial Sector Total", _
                       "Transportation Sector Primary", "Transportation Sector Total", _
                       "Electric Power Sector Primary", "Energy Consumption Balancing Item", _
                       "Grand Consumption Total")
        Set thNode = docobj.createElement("th")
        thNode.Text = i
        trNode.appendChild thNode
    Next i
    
    Set runHeaders = trNode
End Function
