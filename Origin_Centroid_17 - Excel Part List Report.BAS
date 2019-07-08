
'© 2013 Mentor Graphics Corporation.  All Rights Reserved.
 
'This software or file (“Software”) contains trade secret and proprietary information that is the property of Mentor 'Graphics Corporation (“Mentor Graphics”) or its licensors, and is subject to license terms.  If you have a signed license 'agreement with Mentor Graphics or a Mentor Graphics subsidiary for the product(s) with which this Software will be 'used, your use of this Software is subject to the scope of license and the software protection and security provisions of 'that agreement.  If you do not have such a signed agreement, your use is subject to Mentor Graphics’ standard End-'User License Agreement (“EULA”), which may be viewed at http://www.mentor.com/terms_conditions/enduser.
 
'DISCLAIMER:  NOTWITHSTANDING ANYTHING TO THE CONTRARY IN YOUR SIGNED LICENSE AGREEMENT OR THE 'EULA, THIS SOFTWARE IS PROVIDED “AS IS” WITH ALL FAULTS AND WITH:
 
'(A)          NO warranty of any kind, express, implied or statutory, including any implied warranties of merchantability or 'fitness for a particular purpose, which Mentor Graphics disclaims to the maximum extent permitted by applicable law; and
 
'(B)          NO indemnification for infringement of intellectual property rights.
 
'LIMITATION OF LIABILITY:  Because this Software is provided “AS IS”, NEITHER MENTOR GRAPHICS NOR ITS 'LICENSORS SHALL BE LIABLE FOR ANY DAMAGES WHATSOEVER IN CONNECTION WITH THE SOFTWARE OR ITS USE.  'Without limiting the foregoing, in no event will Mentor Graphics or its licensors be liable for indirect, special, incidental, 'or consequential damages (including lost profits or savings) whether based on contract, tort (including negligence), 'strict liability, or any other legal theory, even if Mentor Graphics or its licensors have been advised of the possibility of 'such damages.  THE FOREGOING LIMITATIONS SHALL APPLY TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW.
 
'Unless otherwise agreed in writing, Mentor Graphics has no obligation to support or otherwise maintain Software.

'Array of column names. You can modify it to rename columns
Const Columns = Array("PartType", "RefDes", "PartDecal", "Pins", "Layer", "Orient.", "Origin X", "Origin Y", "Centroid X", "Centroid Y", "SMD", "Glued")

Sub Main
	tempFile = DefaultFilePath & "\temp.txt"
	Open tempFile For Output As #1

	'Output table header
	For i = 0 to UBound(Columns)
		OutCell Columns(i)
	Next
	Print #1
	'Output table rows
	For Each part in ActiveDocument.Components
		OutCell part.PartType
		OutCell part.Name
		OutCell part.Decal
		OutCell part.Pins.Count
		OutCell ActiveDocument.LayerName(part.layer)
		OutCell part.orientation
    		OutCell Format(part.PositionX, "0.000") 'Creates locations based on the decal origin.
    		OutCell Format(part.PositionY, "0.000")
		OutCell Format(part.CenterX, "0.000") 'Creates locations based on the outermost extents
		OutCell Format(part.CenterY, "0.000") 'of pins and part outline making up the decal. 
		OutCell Format(part.IsSMD, "Yes/No")
		OutCell Format(part.Glued, "Yes/No")
		Print #1
	Next part

	Close #1
	ExportToExcel
End Sub

Sub ExportToExcel
	FillClipboard
	Dim xl As Object
	On Error Resume Next
	Set xl =  GetObject(,"Excel.Application")
	On Error GoTo ExcelError	' Enable error trapping.
	If xl Is Nothing Then
		Set xl =  CreateObject("Excel.Application")
	End If
	xl.Visible = True
	xl.Workbooks.Add
	xl.ActiveSheet.Paste
	xl.Range("A1:J1").Font.Bold = True
	xl.Range("A1:J1").NumberFormat = "@"
	xl.ActiveSheet.UsedRange.Columns.AutoFit
	xl.Range("A1").Select
	On Error GoTo 0 ' Disable error trapping. 
	Exit Sub    

ExcelError:
	MsgBox Err.Description, vbExclamation, "Error Running Excel"
	On Error GoTo 0 ' Disable error trapping.    
	Exit Sub
End Sub

Sub OutCell (txt As String)
	Print #1, txt; vbTab;
End Sub

Sub FillClipboard
	' Load whole file to string variable    
	tempFile = DefaultFilePath & "\temp.txt"
	Open tempFile  For Input As #1
	L = LOF(1)
	AllData$ = Input$(L,1)
	Close #1
	'Copy whole data to clipboard
	Clipboard AllData$ 
	Kill tempFile
End Sub
