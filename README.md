Real-Time-Energy-Consumption-Across-a-Power-Plant
=================================================

Public DBOPEN As Long

Sub RevenueTracker()

'Defined
Dim Cn As ADODB.Connection
Dim Server_Name As String
Dim Database_Name As String
Dim User_ID As String
Dim Password As String
Dim SQLStr As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim Uid As String
Dim Pwd As String
Dim numrow As Long
Dim i As Long
' Written by Nicholas Papiri 1/1/2012 - FXXXX XXXD INC.

'Select Dates
COD = Format(Worksheets("Summary").Range("B249").Value, "m/d/yyyy hh:mm:ss")
Startdates = Format(Worksheets("Summary").Range("C3").Value, "m/d/yyyy hh:mm:ss")
enddates = Format(Worksheets("Summary").Range("C4").Value, "m/d/yyyy hh:mm:ss")
Server_Name = "10.7.32.22" ' Enter your server name here
Uid = "Data_Pull" ' enter your user ID here
Password = "Pull_Data" ' Enter your password here
Database_Name = "MIL" ' Enter your database name here

SQLStr = "Select DTTM, RTAC_Device_Cntrs_F1Meter_345kV_KWHDel_Frozen_VALUE as 'F1', RTAC_Device_Cntrs_F2Meter_345kV_KWHDel_Frozen_VALUE as 'F2',RTAC_Device_Cntrs_F3Meter_345kV_KWHDel_Frozen_VALUE as 'F3',RTAC_Device_Cntrs_F4Meter_345kV_KWHDel_Frozen_VALUE as 'F4',RTAC_Device_Cntrs_F5Meter_345kV_KWHDel_Frozen_VALUE as 'F5',RTAC_Device_Cntrs_F6Meter_345kV_KWHDel_Frozen_VALUE as 'F6',RTAC_Device_Cntrs_F7Meter_345kV_KWHDel_Frozen_VALUE as 'F7', RTAC_Device_Cntrs_Bus1Meter_345kV_KWHDel_Frozen_VALUE as 'Buss I', RTAC_Device_Cntrs_Bus2Meter_345kV_KWHDel_Frozen_VALUE as 'Buss II',RTAC_Device_Cntrs_SSMeter_345kV_KWHDel_Frozen_Value as 'SSM', RTAC_Device_Cntrs_IPPMeter_345kV_KWHRec_Frozen_VALUE as 'IPP', RTAC_Device_Cntrs_IPPMeter_345kV_KWHRec_Frozen_VALUE as 'KVARh' from dbo.RTAC_cntrs where DTTM between '" & Startdates & "' and '" & enddates & "' order by DTTM asc"


'Checks the dates to verify if the user selected the wrong values
If Startdates < enddates Then
    If startdate > COD Then
        MsgBox " Please be aware that you have selected a date great than the COD Date"
        End If

DBOPEN = 0
'Open DB
Set Cn = New ADODB.Connection

Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Trusted_Connection=no;Database=" & Database_Name & _
";Uid=" & Uid & ";Pwd=" & Password & ";"


If Cn.State = adStateOpen Then MsgBox " Congrats!The Database has been Opened"
If Cn.State <> adStateOpen Then MsgBox "The Database was not opened, Please Check the Server"

rs.Open SQLStr, Cn, adOpenDynamic


' Dump to spreadsheet
With Worksheets("Phase II Calculation SQL Query").Range("o3:AA3") ' Enter your sheet name and range here
    .ClearContents
    .CopyFromRecordset rs
End With

'Format the cells for the correct representative values ( ie numbers)
Sheets("Phase II Calculation SQL Query").Select
   Range("o3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "mm/dd/yyyy hh:mm"
    Range("P3:AA3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00"
'Return to Original Sheet
    Sheets("Summary").Select
'Take the Delta of the IPP between current timestamp and previous timestamp

' Count the number of records
'If Cn.State = adStateOpen Then
'Sheets("sheet1").Select
'numrow = ActiveSheet.UsedRange.Rows.Count
 
'For i = 2 To numrow
'If Cn.State = adStateOpen Then
'Range("J" & i) = "Values"
'End If
'Next i
    
'End If
rs.Close
Set rs = Nothing
Cn.Close
Set Cn = Nothing

ElseIf Startdates > enddates Then
DBOPEN = 1
MsgBox "The start date provided must be less than the end date. Please select another start date!"
End If

End Sub

Sub CopyData()

If DBOPEN = 0 Then


        
 'This Code Clears the Data From Both Sheets Phase II Calculation and Hourly Summary
    Sheets("Phase II Calculation SQL Query").Select
    Range("o3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    
    Sheets("Phase II Calculation SQL Query").Select
    Range("A6:N6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    

      
    Sheets("Summary").Select
    MsgBox "The Summary page now contains values per your query; please save the document as it is and exit"
    
    ElseIf DBOPEN = 1 Then
       
End If
    
End Sub

Sub Propogate()


 If DBOPEN = 0 Then

 'This Code Drags and autofills all major calculations,which is dependent on the number of rows calculated in the query,
 'in the spread sheet and returns the user back to the interface

    Dim mycount As Long
    Sheets("Phase II Calculation SQL Query").Select
    mycount = Application.CountA(Range("o:o"))

    MsgBox "The total number of entries for this query is: " & mycount&
    Range("A4:N4").Select
    Selection.AutoFill Destination:=Range("A4:N" & mycount)
      ElseIf DBOPEN = 1 Then
      End If
      
End Sub

'This code combines the previous subroutines into a single subroutine and runs them in order
Sub Main()
Call RevenueTracker

Call CopyData
End Sub


