# Dictionary-VBA
Dictionary Class in VBA

https://excelmacromastery.com/vba-dictionary/

https://www.experts-exchange.com/articles/3391/Using-the-Dictionary-Class-in-VBA.html

https://stackoverflow.com/questions/915317/does-vba-have-dictionary-structure

https://github.com/VBA-tools/VBA-Dictionary/blob/master/Dictionary.cls

https://www.experts-exchange.com/articles/3391/Using-the-Dictionary-Class-in-VBA.html

https://excelmacromastery.com/vba-dictionary/

## What Is a Dictionary?
Part of the Microsoft Scripting Runtime (scrrun.dll) library, the Dictionary class allows you to create objects holding an arbitrary number of items, with each item identified by a unique key.  A Dictionary object can hold items of any data type (including other objects, such as other Dictionaries).  A Dictionary's keys can also be any data type except for arrays, although in practice they are almost always either strings or Integer/Long values.  A single Dictionary object can store items of a mix different data types, and use keys of a mix of different data types.

Procedures that create a Dictionary can then:
- Add new items to the Dictionary; 
- Remove items from the Dictionary; 
- Retrieve items from the Dictionary by referring to their associated key values; 
- Change the item associated with a particular key; 
- Retrieve the set of all keys currently in use; 
- Retrieve the count of keys currently in use; and 
- Change a key value, if needed. 

## How a Dictionary Differs from a Collection
VBA developers will recognize a resemblance to the Collection class.  The Collection class is native to the VBA library, and as such is fully integrated into the language.  Thus, no special steps are required to use a Collection object.

Like a Dictionary, when you create a Collection you can then:
- Add an arbitrary number of items to it, of any data type (like Dictionaries, this can include objects, as well as other Collections); 
- Remove items from it; 
- Retrieve items from it; and 
- Return a count of items in the Collection. 

However, Collections and Dictionaries have the following differences:

- For Dictionaries, keys are mandatory and always unique to that Dictionary.  In a Collection, while keys must be unique, they            are also optional. 
- In a Dictionary, an item can only be returned in reference to its key.  In a Collection, and item can be returned in reference to its key, or in reference to its index value (i.e., ordinal position within the Collection, starting with 1). 
- With a Dictionary, the key can take any data type; for string keys, by default a Dictionary is case sensitive, but by changing the CompareMode property it can be made case insensitive.  In a Collection, keys are always strings, and always case insensitive.  (See Example #2: Distinct Values with Case-Sensitive Keys) 
- With a Dictionary, there is an Exists method to test for the existence of a particular key (and thus of the existence of the item associated with that key).  Collections have no similar test; instead, you must attempt to retrieve a value from the Collection, and handle the resulting error if the key is not found (see the entry for the Exists method in section Dictionary Properties and Methods below). 
- A Dictionary's items and keys are always accessible and retrievable to the developer.  A Collection's items are accessible and retrievable, but its keys are not.  Thus, for any operation in which retrieval of the keys is as important as retrieval of the items associated with those keys, a Dictionary object will enable a cleaner implementation than a Collection will. 
- The Dictionary's Item property is read/write, and thus it allows you to change the item associated with a particular key.  A Collection's Item property is read-only, and so you cannot reassign the item associated with a specified key: you must instead remove that item from the Collection, and then add in the new item. 
- A Dictionary allows you to change a particular key value.  (This is distinct from changing the value associated with a particular key.)  A Collection will not allow you to do this; the nearest you could come is to remove the item using the former key value, and then to add the item back using the new key value. 
- A Dictionary allows you to remove all items in a single step without destroying the Dictionary itself.  With a Collection, you would have to either remove each item in turn, or destroy and then recreate the Collection object. 
- Both Dictionaries and Collections support enumeration via For...Each...Next.  However, while for a Collection this enumerates the items, for a Dictionary this will enumerate the keys.  Thus, to use For...Each...Next to enumerate the items in a Dictionary: 
- A Dictionary supports implicit adding of an item using the Item property.  With Collections, items must be explicitly added. 


## Dictionary Properties and Methods

The Dictionary class has four properties and six methods, as discussed below.

## Examples

```
Sub MakeTheList()
    
    Dim dic As Object
    Dim dic2 As Object
    Dim Contents As Variant
    Dim ParentKeys As Variant
    Dim ChildKeys As Variant
    Dim r As Long, r2 As Long
    Dim LastR As Long
    Dim WriteStr As String
    
    ' Create "parent" Dictionary.  Each key in the parent Dictionary will be a disntict
    ' Code value, and each item will be a "child" dictionary.  For these "children"
    ' Dictionaries, each key will be a distinct Product value, and each item will be the
    ' sum of the Quantity column for that Code - Product combination
    
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    
    ' Dump contents of worksheet into array
    
    With ThisWorkbook.Worksheets("Data2")
        LastR = .Cells(.Rows.Count, 1).End(xlUp).row
        Contents = .Range("a2:c" & LastR).Value
    End With
        
    ' Loop through the array
    
    For r = 1 To UBound(Contents, 1)
        
        ' If the current code matches a key in the parent Dictionary, then set dic2 equal
        ' to the "child" Dictionary for that key
        
        If dic.Exists(Contents(r, 1)) Then
            Set dic2 = dic.Item(Contents(r, 1))
            
            ' If the current Product matches a key in the child Dictionary, then set the
            ' item for that key to the value of the item now plus the value of the current
            ' Quantity
            
            If dic2.Exists(Contents(r, 2)) Then
                dic2.Item(Contents(r, 2)) = dic2.Item(Contents(r, 2)) + Contents(r, 3)
            
            ' If the current Product does not match a key in the child Dictionary, then set
            ' add the key, with item being the amount of the current Quantity
            
            Else
                dic2.Add Contents(r, 2), Contents(r, 3)
            End If
        
        ' If the current code does not match a key in the parent Dictionary, then instantiate
        ' dic2 as a new Dictionary, and add an item (Quantity) using the current Product as
        ' the Key.  Then, add that child Dictionary as an item in the parent Dictionary, using
        ' the current Code as the key
        
        Else
            Set dic2 = CreateObject("Scripting.Dictionary")
            dic2.CompareMode = vbTextCompare
            dic2.Add Contents(r, 2), Contents(r, 3)
            dic.Add Contents(r, 1), dic2
        End If
    Next
    
    ' Add a new worksheet for the results
    
    Worksheets.Add
    [a1:b1].Value = Array("Code", "Product - Qty")
    
    ' Dump the keys of the parent Dictionary in an array
    
    ParentKeys = dic.Keys
    
    ' Write the parent Dictionary's keys (i.e., the distinct Code values) to the worksheet
    
    [a2].Resize(UBound(ParentKeys) + 1, 1).Value = Application.Transpose(ParentKeys)
    
    ' Loop through the parent keys and retrieve each child Dictionary in turn
    
    For r = 0 To UBound(ParentKeys)
        Set dic2 = dic.Item(ParentKeys(r))
        
        ' Dump keys of child Dictionary into array and initialize WriteStr variable (which will
        ' hold concatenated products and summed Quantities
        
        ChildKeys = dic2.Keys
        WriteStr = ""
        
        ' Loop through child keys and retrieve summed Quantity value for that key.  Build both
        ' of these into the WriteStr variable.  Recall that Excel uses linefeed (ANSI 10) for
        ' in-cell line breaks
        
        For r2 = 0 To dic2.Count - 1
            WriteStr = WriteStr & Chr(10) & ChildKeys(r2) & " - " & dic2.Item(ChildKeys(r2))
        Next
        
        ' Trim leading linefeed
        
        WriteStr = Mid(WriteStr, 2)
        
        ' Write concatenated list to worksheet
        
        Cells(r + 2, 2) = WriteStr
    Next
    
    ' Sort and format return values
    
    [a1].Sort Key1:=[a1], Order1:=xlAscending, Header:=xlYes
    With [b:b]
        .ColumnWidth = 40
        .WrapText = True
    End With
    Columns.AutoFit
    Rows.AutoFit
    
    ' Destroy object variables
    
    Set dic2 = Nothing
    Set dic = Nothing
    
    MsgBox "Done"
    
End Sub
```

## Dictionary para suma acumulativa, dejando valor total en columna E con "1"

```
Sub Funcionando_Suma_Acumulada()
    'https://excelmacromastery.com/vba-dictionary/
    Dim sh As Worksheet
    Dim lastRow As Integer
    Dim dict As Object
    Dim i As Integer
    Dim Contents As Variant
    Set sh = Worksheets("Hoja1")
    lastRow = sh.Range("A" & Rows.Count).End(xlUp).row 'conteo de columna
    
    Set dict = CreateObject("Scripting.Dictionary") 'Create(late binding)
    dict.CompareMode = vbTextCompare 'Make key non case sensitive (the dictionary must be empty)
    Contents = sh.Range("A2:B" & lastRow).Value
    
    For i = 1 To UBound(Contents, 1)
        If Not dict.Exists(Contents(i, 1)) Then 'columna de conteo columna 1
            dict.Add Contents(i, 1), Contents(i, 2)
            'sh.Cells(i, 3).Value = dict(sh.Cells(i, 1).Value) 'columna donde deja = columna de conteo columna 1
        Else
            dict.Item(Contents(i, 1)) = dict.Item(Contents(i, 1)) + Contents(i, 2) 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            'sh.Cells(i, 3).Value = dict.Item(Contents(i, 1)) 'columna donde deja = columna de conteo columna 1
        End If
    Next i
     
    For i = 2 To lastRow
        If Cells(i, 5).Value = "1" Then
            Cells(i, 6) = dict(Cells(i, 1).Value)
        End If
    Next i
    'Arreglar columnas
    Columns.AutoFit
    Rows.AutoFit
    'Destroy object variables
    Set dict = Nothing
End Sub
```
Resultado

![Dictionary](https://user-images.githubusercontent.com/17385297/56093380-1318ab80-5e96-11e9-95bd-f2174d778441.PNG)



## Dictionary para conteo de lineas, enumarando las lineas segun Key

```
Sub countPO2()

Dim ws As Worksheet
Dim lastRow As Long, x As Long
Dim items As Object

Application.ScreenUpdating = False
  
Set ws = Worksheets("Base")
Cells(5, Ultimate_Column + 6).Value = "lineas" 'Donde va el titulo
            Cells(5, Ultimate_Column + 6).Font.Bold = True 'Donde va el titulo en negrita
    
lastRow = ws.Range("E" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set items = CreateObject("Scripting.Dictionary")
    For x = 6 To lastRow
        If Not items.Exists(ws.Cells(x, 1).Value) Then 'columna de conteo columna 1
            items.Add ws.Cells(x, 1).Value, 1 'columna de conteo columna 1
            ws.Cells(x, Ultimate_Column + 6).Value = items(ws.Cells(x, 1).Value) 'columna donde deja = columna de conteo columna 1
        Else
            items(ws.Cells(x, 1).Value) = items(ws.Cells(x, 1).Value) + 1 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            ws.Cells(x, Ultimate_Column + 6).Value = items(ws.Cells(x, 1).Value) 'columna donde deja = columna de conteo columna 1
        End If
    Next x
End Sub
```
Resultado

![Dictionary 2](https://user-images.githubusercontent.com/17385297/56093473-3bed7080-5e97-11e9-852e-8a55ccc406c3.PNG)


## Dictionary para mensaje acumulativo segun email y cantidad de linea. Seleccionar el mensaje mas completo segun email

```
Sub Funcionando_Suma_Acumulada2()
    'https://excelmacromastery.com/vba-dictionary/
    Dim sh As Worksheet
    Dim lastRow As Integer
    Dim dict As Object
    Dim i As Integer
    Dim Contents As Variant

    Set sh = Worksheets("Hoja1")
    lastRow = sh.Range("B" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set dict = CreateObject("Scripting.Dictionary") 'Create(late binding)
    dict.CompareMode = vbTextCompare 'Make key non case sensitive (the dictionary must be empty)
    Contents = sh.Range("B1:G" & lastRow).Value 'Campo donde esta la informacion
    
    For i = 2 To UBound(Contents, 1)
        If Not dict.exists(Contents(i, 1)) Then 'La columna 3, sera la columna comparativa
            dict.Add Contents(i, 1), "N° Ticket " & Contents(i, 2) & " - " & "N° Solicitud " & Contents(i, 3) & " - " & "Estado Solicitud " & Contents(i, 4) & " - " & "Estado Presupuesto " & Contents(i, 5) & " - " & "SLA " & Contents(i, 6) & " dias"
        Else
            dict(Contents(i, 1)) = dict(Contents(i, 1)) & vbNewLine & "N° Ticket " & Contents(i, 2) & " - " & "N° Solicitud " & Contents(i, 3) & " - " & "Estado Solicitud " & Contents(i, 4) & " - " & "Estado Presupuesto " & Contents(i, 5) & " - " & "SLA " & Contents(i, 6) & " dias"
        End If
    
    Cells(i, 8) = dict(Contents(i, 1))
    'Debug.Print (dict(Contents(i, 1)))
    'Debug.Print ("---------------------------------------------------------------")
    Next i

    'Destroy object variables
    Set dict = Nothing
    
    'Debug.Print ("----------TERMINO----------")

Call Envio_masivo_email
End Sub
```

Resultado

![Macro email](https://user-images.githubusercontent.com/17385297/63657731-06c12b00-c773-11e9-8f83-69b5aa8a7475.JPG)

## Seleccionar email con la mayor cantidad de linea y el mensaje mas completo.
## Envio masivo de email

```
Sub Envio_masivo_email()
'https://stackoverflow.com/questions/54370800/find-max-min-value-in-dictionary-values-in-vba/54371824
Dim ws As Worksheet
Dim lastRow As Long, x As Long
Dim dict As Object
Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String
Dim i As Integer


Application.ScreenUpdating = False
  
Set ws = Worksheets("Hoja1")
Cells(1, 9).Value = "Conteo de email" 'Donde va el titulo
            Cells(1, 9).Font.Bold = True 'Donde va el titulo en negrita
    
lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set dict = CreateObject("Scripting.Dictionary")
    For x = 2 To lastRow
        If Not dict.exists(ws.Cells(x, 2).Value) Then 'columna de conteo columna 1
            dict.Add ws.Cells(x, 2).Value, 1 'columna de conteo columna 1
            ws.Cells(x, 9).Value = dict(ws.Cells(x, 2).Value) 'columna donde deja = columna de conteo columna 1
        Else
            dict(ws.Cells(x, 2).Value) = dict(ws.Cells(x, 2).Value) + 1 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            ws.Cells(x, 9).Value = dict(ws.Cells(x, 2).Value) 'columna donde deja = columna de conteo columna 1
        End If
    Next x
        
    max = Application.max(dict.items)
    min = Application.min(dict.items)
    
'Debug.Print key

    For Each key In dict
        For x = 2 To lastRow
        If Cells(x, 2) = key And Cells(x, 9) = dict(key) Then
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.CreateItem(0)
            
            strbody = "Estimados" & vbNewLine & vbNewLine & _
                        "Las siguientes solicitudes: " & vbNewLine & vbNewLine & _
                        Cells(x, 8).Text & vbNewLine & vbNewLine & _
                        "Gracias"
            On Error Resume Next
            With OutMail
                '.To = "matias.vidal@serviciosandinos.net"
                .To = Cells(x, 2).Text
                .CC = "jose.fuentes@serviciosandinos.net"
                .BCC = ""
                '.Subject = Worksheets("Hoja1").Cells(i, 3).Value
                .Subject = "Solicitudes Pendientes SSC"
                .Body = strbody
                'You can add a file like this
                '.Attachments.Add ("C:\test.txt")
                .Importance = 2
                .Send   'or use .Display
            End With
            On Error GoTo 0
            Set OutMail = Nothing
            Set OutApp = Nothing
        End If
        Next x
    Debug.Print key
    Debug.Print dict(key)
    Debug.Print lastRow
    Next


'Range("G1:H10000").Clear
Columns([9]).EntireColumn.Delete
Columns([8]).EntireColumn.Delete
Columns.AutoFit
Rows.AutoFit
End Sub

```
## Dictionary para sumar los valores asociados a un Key.

Contensts es la matriz que comienza con la llave y llega hasta el value.
```
Contents = sh.Range("A1:C" & lastRow).Value 'Campo donde esta la informacion
```
UBound(Contents, 1). Indica la matriz donde se busca.

```
For i = 2 To UBound(Contents, 1)
```

En caso que no exista la Key, asociar la Key Contents(i, 1) con item Contents(i, 3)
En caso que existe agregar con una salto de linea vbNewLine
```

If Not dict.exists(Contents(i, 1)) Then 'La columna H, sera la columna comparativa
    dict.Add Contents(i, 1), Contents(i, 3)
Else
    dict.Item(Contents(i, 1)) = dict.Item(Contents(i, 1)) & vbNewLine & Contents(i, 3)
End If
Next i
```
Codigo completo

```
Sub Dictionary_cuenta_contable_proveedor()
    'https://excelmacromastery.com/vba-dictionary/
    Dim sh As Worksheet
    Dim sh2 As Worksheet
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim dict As Object
    Dim i As Long
    Dim Contents As Variant

    Worksheets("Status_data_total").Activate
    Set sh = Worksheets("Status_data_total")
    lastRow = sh.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set dict = CreateObject("Scripting.Dictionary") 'Create(late binding)
    dict.CompareMode = vbTextCompare 'Make key non case sensitive (the dictionary must be empty)
    Contents = sh.Range("A1:C" & lastRow).Value 'Campo donde esta la informacion
    
    For i = 2 To UBound(Contents, 1)
            If Not dict.exists(Contents(i, 1)) Then 'La columna H, sera la columna comparativa
            dict.Add Contents(i, 1), Contents(i, 3)
            Else
            'If dict.Item(Contents(i, 1)) = Contents(i, 69) Then
                'dict.Item(Contents(i, 1)) = dict.Item(Contents(i, 1))
            'Else
                dict.Item(Contents(i, 1)) = dict.Item(Contents(i, 1)) & vbNewLine & Contents(i, 3)
            'End If
            End If
    Next i

        'For Each Item In dict.Item
        '    Debug.Print dict.Item
        'Next Item
    Worksheets("EXCEPCIONES(total)").Activate
    Set sh2 = Worksheets("EXCEPCIONES(total)")
    lastRow2 = sh2.Range("A" & Rows.Count).End(xlUp).Row 'conteo de columna
    'Contents = sh2.Range("A1:L" & lastRow).Value 'Campo donde esta la informacion
    sh2.Cells(1, 7) = "Cuenta Contable"
    For i = 2 To lastRow2
        If Not dict.exists(sh2.Cells(i, 2).Value) Then
            sh2.Cells(i, 7) = "No cuenta contable"
        Else
            sh2.Cells(i, 7) = dict(sh2.Cells(i, 2).Value)
        End If
    Next i
    Columns.AutoFit
    Rows.AutoFit
    'Destroy object variables
    Set dict = Nothing
    Debug.Print ("----------TERMINO-Dictionary_cuenta_contable_proveedor---------")
'Call Limpiar_String
End Sub
```
