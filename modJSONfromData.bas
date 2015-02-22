Attribute VB_Name = "modJSONfromData"
Option Explicit
' JSON serialisation method inspired by amadeus answer on
' http://stackoverflow.com/questions/2782076/is-there-a-json-parser-for-vb6-vba
' though actually we are only using that for validation
Private JSON As Object
Private ie As Object

Public Sub initJson()
    Dim html As String

    html = "<!DOCTYPE html><head><script>" & _
    "Object.prototype.getItem=function( key ) { return this[key] }; " & _
    "Object.prototype.setItem=function( key, value ) { this[key]=value }; " & _
    "Object.prototype.getKeys=function( dummy ) { keys=[]; for (var key in this) if (typeof(this[key]) !== 'function') keys.push(key); return keys; }; " & _
    "window.onload = function() { " & _
    "document.body.parse = function(json) { return JSON.parse(json); }; " & _
    "document.body.stringify = function(obj, space) { return JSON.stringify(obj, null, space); }" & _
    "}" & _
    "</script></head><html><body id='JSONElem'></body></html>"
    If ie Is Nothing Then
      Set ie = CreateObject("InternetExplorer.Application")
    End If
    If JSON Is Nothing Then
      With ie
          .navigate "about:blank"
          Do While .Busy: DoEvents: Loop
          Do While .ReadyState <> 4: DoEvents: Loop
          .Visible = False
          .Document.Write html
          .Document.Close
      End With
  
      ' This is the body element, we call it JSON:)
      Set JSON = ie.Document.getElementById("JSONElem")
    End If
End Sub

Private Sub closeJSON()
    ie.Quit
    Set ie = Nothing
    Set JSON = Nothing
End Sub

Function JSONfromRecordset(rs As Recordset, bPerformValidation As Boolean) As String
  Dim sJSON As String
  'Call initJson  ' rejected approach
  Dim jsObj As Object
  Dim sTmpJSON As String
  Dim i As Long
  Dim fieldVal As Variant
  Dim sRelatedSQL As String, sDomain As String, sCondition As String, sPKFieldName As String, sPKVal As String, sFKFieldName As String
  
On Error GoTo err_JSONfromRecordset

  Do Until rs.EOF
    'Set jsObj = JSON.Parse("{}")
    ' rejected that appproach as it cannot cope with adding the array of related records to the object
    ' so we are building the record JSON as a string
    sTmpJSON = "{"
    For i = 0 To rs.Fields.Count - 1
      If i > 0 Then sTmpJSON = sTmpJSON & ","
      fieldVal = rs(i)
      ' check for our convention that fetches related values from a related table or query
      ' example: GetRelated(Cats,OwnerID,ID) where Cats is the related table, OwnerID is the FK, ID is the key in this table
      If Left(fieldVal, 11) = "GetRelated(" Then
        sDomain = Mid(fieldVal, 12, InStr(fieldVal, ",") - 12)
        sFKFieldName = Mid(fieldVal, 12 + Len(sDomain) + 1, InStrRev(fieldVal, ",") - (12 + Len(sDomain) + 1))
        sPKFieldName = Mid(fieldVal, InStrRev(fieldVal, ",") + 1, Len(fieldVal) - InStrRev(fieldVal, ",") - 1)
        ' safe but not very efficient coercing all values to strings - TODO improve this
        sPKVal = CStr(Nz(rs(sPKFieldName)))
        sCondition = " WHERE CStr([" & sFKFieldName & "]) = '" & sPKVal & "'"
        sRelatedSQL = "SELECT * FROM [" & sDomain & "] " & sCondition
        If sPKVal > "" Then ' if empty, there can't be any related records
          fieldVal = JSONfromData(sRelatedSQL, bPerformValidation)
          sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & fieldVal
        End If
      Else
        Select Case rs.Fields(i).Type
          Case dbInteger, dbLong, dbSingle, dbDouble, dbBigInt, dbNumeric, dbDecimal, dbFloat ' numbers
            ' bizarrely this doesn't work: Call jsObj.setItem(rs.Fields(i).Name, rs(i))
            ' so we collect the value first, then it works
            'Call jsObj.setItem(rs.Fields(i).Name, fieldVal) ' rejected approach
            sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & IIf(IsNull(fieldVal), "null", fieldVal)
          Case dbCurrency
            'Call jsObj.setItem(rs.Fields(i).Name, fieldVal)
          Case dbText, dbMemo, dbGUID, dbChar
            'Call jsObj.setItem(rs.Fields(i).Name, "" & fieldVal & "") ' rejected approach
            sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & IIf(IsNull(fieldVal), "null", """" & escapeJSON(Nz(fieldVal)) & """")
          Case dbDate
            'Call jsObj.setItem(rs.Fields(i).Name, "" & Format(fieldVal, "yyyy-mm-dd\Thh:nn:ss") & "") ' rejected approach
            sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & IIf(IsNull(fieldVal), "null", """" & Format(fieldVal, "yyyy-mm-dd\Thh:nn:ss") & """")
          Case dbBoolean
            'Call jsObj.setItem(rs.Fields(i).Name, fieldVal) ' rejected approach
            sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & IIf(IsNull(fieldVal), "null", LCase(fieldVal))
          Case Else
            'Call jsObj.setItem(rs.Fields(i).Name, fieldVal) ' rejected approach
            sTmpJSON = sTmpJSON & """" & rs.Fields(i).Name & """: " & fieldVal
        End Select
      End If
    Next
    'sJSON = sJSON & "," & JSON.stringify(jsObj, 2) ' rejected approach
    sJSON = sJSON & "," & sTmpJSON & "}"
    'Set jsObj = Nothing ' rejected approach
    rs.MoveNext
  Loop
  ' remove leading comma and wrap as array of records
  sJSON = "[" & Mid(sJSON, 2) & "]"
  ' check and clean
  If bPerformValidation Then
    Call initJson
    Set jsObj = JSON.Parse(sJSON)
    sJSON = JSON.stringify(jsObj, 2)
    Call closeJSON ' this makes performance MUCH slower if we are calling related records, hence making validation optional
  End If

exit_JSONfromRecordset:
  JSONfromRecordset = sJSON
  Exit Function
  
err_JSONfromRecordset:
  sJSON = "Error. " & Err.Description
  Resume exit_JSONfromRecordset
  
End Function
Private Function escapeJSON(sJSON As String) As String
'http://json.org/
  
  sJSON = Replace(sJSON, "\", "\\") ' do this one first of course
  sJSON = Replace(sJSON, vbBack, "\n")
  sJSON = Replace(sJSON, vbLf, "\n")
  sJSON = Replace(sJSON, vbCr, "\r")
  sJSON = Replace(sJSON, vbTab, "\t")
  sJSON = Replace(sJSON, vbFormFeed, "\f")
  sJSON = Replace(sJSON, """", "\""")
  'sJSON = Replace(sJSON, "/", "\/") ' there appears to be no need for this despite spec
  escapeJSON = sJSON
End Function

Function JSONfromData(sDataSource As String, Optional bPerformValidation As Boolean = False) As String
  'sDataSource can be a table or query name, or a valid SQL SELECT string
  Dim sResult As String
  Dim rs As Recordset

On Error GoTo err_JSONfromData

  Set rs = CurrentDb.OpenRecordset(sDataSource, dbOpenDynaset, dbSeeChanges)
  sResult = JSONfromRecordset(rs, bPerformValidation)
  rs.Close
  Set rs = Nothing

exit_JSONfromData:
  JSONfromData = sResult
  Exit Function

err_JSONfromData:
  sResult = "Error opening " & sDataSource & ". " & Err.Description
  Resume exit_JSONfromData

End Function

