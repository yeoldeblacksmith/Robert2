<%
Class ArrayList
	Private CAPACITY
	Private m_arrData()
	Private m_Count
	Private m_DefaultComparer
	Private Sub Class_Initialize
		CAPACITY=100
		Redim m_arrData(CAPACITY-1)
		m_Count=0
		m_DefaultComparer=""
	End Sub
	
	Public Property Get Count
		Count=m_Count
	End Property
	
	Public Property Get DefaultComparer
		DefaultComparer=m_DefaultComparer
	End Property
	
	Public Property Let DefaultComparer(Value)
		m_DefaultComparer=Value
	End Property

	Public Sub Add(value)
		If m_Count>UBound(m_arrData) Then
			Call MakeMorePlace
		End If
		If IsObject(value) Then
			Set m_arrData(m_Count)=value
		Else  
			m_arrData(m_Count)=value
		End If
		m_Count=m_Count+1
	End Sub

    public function Contains(value)
        dim found : found = false
        dim index : index = 0
        
        do while index < Count and found = false
            found = cbool(Item(index) = value)
            index = index + 1
        loop
        
        Contains = found
    end function

    Public Sub Remove(value)
		Dim x
		For x=0 To m_Count-1
			If Compare(m_arrData(x), value, m_DefaultComparer)=0 Then
				m_arrData(x)=Empty
			End If
		Next
		Call RemoveEmptyItems
	End Sub
	
	Public Default Property Get Item(index)
		If Not(IsNumeric(index)) Then
			Err.Raise 15000, "ArrayList", "invalid parameter given to Item: not numeric"
		End If
		If (index<0) Or (index>=m_Count) Then
			Err.Raise 15001, "ArrayList", "index out of range: ("&index&") can't be less than zero or greater than count"
		End If
		If IsObject(m_arrData(index)) Then
			Set Item=m_arrData(index)
		Else  
			Item=m_arrData(index)
		End If
	End Property
	
	Public Function ToArray()
		Dim result(), x
		If m_Count=0 Then
			ToArray=Array() 'empty array
			Exit Function
		End If
		ReDim result(m_Count-1)
		For x=0 To m_Count-1
			If IsObject(m_arrData(x)) Then
				Set result(x)=m_arrData(x)
			Else  
				result(x)=m_arrData(x)
			End If
		Next
		ToArray=result
	End Function
	
    public function ToJSON()
        dim output

        output = "["
    
        for i = 0 to Count - 1
            if i > 0 then output = output & ","
            output = output & (new JSON)("item", Item(i), false)
        next    

        output = output & "]"

        ToJSON = output
    end function

	Public Sub Sort(strComparer)
		Dim x, y, maxValue
		Dim maxIndex, curValue, tempValue
		
		For x=0 To m_Count-1
			maxIndex=x
			If IsObject(m_arrData(x)) Then
				Set maxValue=m_arrData(x)
			Else  
				maxValue=m_arrData(x)
			End If
			For y=x+1 To m_Count-1
				If IsObject(m_arrData(y)) Then
					Set curValue=m_arrData(y)
				Else  
					curValue=m_arrData(y)
				End If
				If Compare(curValue, maxValue, strComparer)>0 Then
					If IsObject(curValue) Then
						Set maxValue=curValue
					Else  
						maxValue=curValue
					End If
					maxIndex=y
				End If
			Next
			If maxIndex<>x Then
				If IsObject(m_arrData(maxIndex)) Then
					Set tempValue=m_arrData(maxIndex)
				Else  
					tempValue=m_arrData(maxIndex)
				End If
				If IsObject(m_arrData(x)) Then
					Set m_arrData(maxIndex)=m_arrData(x)
				Else  
					m_arrData(maxIndex)=m_arrData(x)
				End If
				If IsObject(tempValue) Then
					Set m_arrData(x)=tempValue
				Else  
					m_arrData(x)=tempValue
				End If
			End If
		Next
	End Sub
	
	Private Sub MakeMorePlace
		ReDim Preserve m_arrData(UBound(m_arrData)+CAPACITY)
	End Sub
	
	Private Sub RemoveEmptyItems
		Dim blnContinue, x, delIndex
		blnContinue=True
		Do Until blnContinue=False
			delIndex=-1
			For x=0 To m_Count-1
				If IsEmpty(m_arrData(x)) Then
					delIndex=x
				End If
			Next
			If delIndex>=0 Then
				For x=delIndex To m_Count-2
					If IsObject(m_arrData(x+1)) Then
						Set m_arrData(x)=m_arrData(x+1)
					Else  
						m_arrData(x)=m_arrData(x+1)
					End If
				Next
				m_arrData(m_Count-1)=Empty
				m_Count=m_Count-1
			End If
			blnContinue=(delIndex >= 0)
		Loop
	End Sub
	
	Private Function Compare(item1, item2, strComparer)
		'look for custom comparer:
		If Len(strComparer)>0 Then
			On Error Resume Next
				Compare=Eval(strComparer&"(item1, item2)")
				If Err.Number<>0 Then
					On Error Goto 0
					Err.Raise 15010, "ArrayList", "Can't compare items: invalid comparer or error on custom comparison: "&Err.Description&" ("&strComparer&")"
				End If
			On Error Goto 0
			Exit Function
		End If
		
		'no default comparer, use built in comparison, if possible:
		If IsObject(item1) Or IsObject(item2) Then
			Err.Raise 15002, "ArrayList", "Can't compare items: no custom comparer is defined and complex items found."
		End If
		'If TypeName(item1)=TypeName(item2) Then
		If item1=item2 Then
			Compare=0
		Else  
			If item1>item2 Then
				Compare=1
			Else  
				Compare=-1
			End If
		End If
		'End If
	End Function

	Private Sub Class_Terminate
		Dim x
		For x=0 To UBound(m_arrData)
			If IsObject(m_arrData(x)) Then
				Set m_arrData(x)=Nothing
			End If
		Next
		Erase m_arrData
	End Sub
End Class

'must return zero if items are equal, positive number if item1 bigger then item2 and negative number if item2 is bigger than item1.
Function MyComparer(item1, item2)
	If TypeName(item1)="String" And TypeName(item2)="String" Then
		MyComparer=StrComp(item1, item2)
		Exit Function
	End If
	If TypeName(item1)="Integer" And TypeName(item2)="Integer" Then
		MyComparer=CLng(item1)-CLng(item2)
		Exit Function
	End If
	If TypeName(item1)="String" And TypeName(item2)="Integer" Then
		If IsNumeric(item1) Then
			MyComparer=CLng(item1)-CLng(item2)
			Exit Function
		End If
		MyComparer=1
		Exit Function
	End If
	If TypeName(item1)="Integer" And TypeName(item2)="String" Then
		If IsNumeric(item2) Then
			MyComparer=CLng(item1)-CLng(item2)
			Exit Function
		End If
		MyComparer=-1
		Exit Function
	End If
	MyComparer=0
End Function

Function NameComparer(file1, file2)
	NameComparer=StrComp(file1.Name, file2.Name, 1)
End Function

Function SizeComparer(file1, file2)
	SizeComparer=file1.Size-file2.Size
End Function

Function DateCreatedComparer(file1, file2)
	DateCreatedComparer=DateDiff("s", file2.DateCreated, file1.DateCreated)
End Function

Function DateLastModifiedComparer(file1, file2)
	DateLastModifiedComparer=DateDiff("s", file2.DateLastModified, file1.DateLastModified)
End Function

Function ExtensionComparer(file1, file2)
	Dim strExtension1, strExtension2, arrTemp
	arrTemp=Split(file1.Name, ".")
	strExtension1=arrTemp(UBound(arrTemp))
	Erase arrTemp
	arrTemp=Split(file2.Name, ".")
	strExtension2=arrTemp(UBound(arrTemp))
	Erase arrTemp
	ExtensionComparer=StrComp(strExtension1, strExtension2, 1)
End Function
%>