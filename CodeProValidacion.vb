Private Sub TextBox1_Change()
End Sub
Private Sub btnArchivo1_Click()
Dim archivo1 As FileDialog
    Set archivo1 = Application.FileDialog(msoFileDialogFilePicker)
    
    archivo1.AllowMultiSelect = False
    archivo1.Title = "Seleccionar el archivo el cual se va a validar las incapacidades"
    
    If archivo1.Show = -1 Then
        UserFormValidator.TextBoxArchivo1.Value = GetFileNameFromPath(archivo1.SelectedItems(1))
    End If
End Sub
Private Sub btnArchivo2_Click()
    Dim archivo2 As FileDialog
    Set archivo2 = Application.FileDialog(msoFileDialogFilePicker)
    
    archivo2.AllowMultiSelect = False
    archivo2.Title = "Seleccionar el archivo final de incapacidades donde se Guardaran"
    
    If archivo2.Show = -1 Then
       UserFormValidator.TextBoxArchivo2.Value = GetFileNameFromPath(archivo2.SelectedItems(1))
    End If
End Sub
Function GetFileNameFromPath(ByVal fullPath As String) As String
    Dim parts() As String
    parts = Split(fullPath, Application.PathSeparator)
    GetFileNameFromPath = parts(UBound(parts))
End Function
Private Sub btnValidar_Click()
    If UserFormValidator.TextBoxArchivo1.Value <> "" And UserFormValidator.TextBoxArchivo2.Value <> "" Then
        ValidarIncapacidades
    Else
        MsgBox "Por favor, carga ambos archivos antes de validar.", vbExclamation, "Advertencia"
    End If
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

    ' Archivo 1
    UserFormValidator.TextBoxArchivo1.Text = "Favor cargar un archivo"
    UserFormValidator.TextBoxArchivo1.Enabled = False
    ' Archivo 2
    UserFormValidator.TextBoxArchivo2.Text = "Favor cargar un archivo"
    UserFormValidator.TextBoxArchivo2.Enabled = False

End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub ValidarIncapacidades()
    Dim archive1 As Worksheet
    Dim archive2 As Worksheet
    Dim endArchive1 As Long
    Dim endArchive2 As Long
    Dim contArchive2 As Long ' Para rastrear el último registro agregado
    Dim id As Long
    Dim i As Long
    Dim j As Long
    
    ' Abrir los archivos
    Set archive1 = Workbooks.Open(UserFormValidator.TextBoxArchivo1.Value).Sheets(1)
    Set archive2 = Workbooks.Open(UserFormValidator.TextBoxArchivo2.Value).Sheets(1)
    
    ' Encontrar la última fila con datos en cada archivo
    endArchive1 = archive1.Cells(archive1.Rows.Count, "B").End(xlUp).Row
    endArchive2 = archive2.Cells(archive2.Rows.Count, "A").End(xlUp).Row
    contArchive2 = endArchive2 ' Inicializar con la última fila actual del archivo 2
    
    Dim addedRecords As Collection
    Set addedRecords = New Collection

    For i = 2 To endArchive1
        id = archive1.Cells(i, "B").Value
        Dim idAndDates As String
        idAndDates = id & "_" & archive1.Cells(i, "F").Value & "_" & archive1.Cells(i, "H").Value
        
        ' Verificar si el registro ya se agregó a archive2
        If Not Contains(addedRecords, idAndDates) Then
            Dim foundMatch As Boolean
            foundMatch = False
            
            For j = 2 To endArchive2
                If archive2.Cells(j, "A").Value = id Then
                    If archive1.Cells(i, "F").Value = archive2.Cells(j, "H").Value Then
                        foundMatch = True
                        Exit For
                    End If
                End If
            Next j
            
            If Not foundMatch Then
                contArchive2 = contArchive2 + 1
                archive2.Cells(contArchive2, "A").Value = id
                archive2.Cells(contArchive2, "C").Value = archive1.Cells(i, "D").Value
                archive2.Cells(contArchive2, "D").Value = archive1.Cells(i, "E").Value
                archive2.Cells(contArchive2, "E").Value = archive1.Cells(i, "D").Value & " " & archive1.Cells(i, "E").Value
                archive2.Cells(contArchive2, "H").Value = archive1.Cells(i, "F").Value
                archive2.Cells(contArchive2, "I").Value = archive1.Cells(i, "G").Value
                archive2.Cells(contArchive2, "J").Value = archive1.Cells(i, "H").Value
                archive2.Cells(contArchive2, "K").Value = archive1.Cells(i, "K").Value
                archive2.Cells(contArchive2, "S").Value = archive1.Cells(i, "N").Value
                
                ' Agregar a la colección de registros agregados
                addedRecords.Add idAndDates
            End If
        End If
NextIteration:
    Next i


    ' Alinear celdas
    archive2.Range("A1:S1").HorizontalAlignment = xlCenter
    archive2.Range("A2:S" & contArchive2).HorizontalAlignment = xlLeft
    archive2.Range("A1:S" & contArchive2).VerticalAlignment = xlCenter

    ' Cerrar los archivos y el formulario
    archive1.Parent.Close False
    archive2.Parent.Close True
    Unload Me
    Application.Quit
    ThisWorkbook.Close SaveChanges:=False

    MsgBox "Todo Validado de forma Correcta", vbInformation
End Sub

Function Contains(col As Collection, item As String) As Boolean
    On Error Resume Next
    Contains = col(item) = item
    On Error GoTo 0
End Function
'creo el link de visita a mi perfil
Private Sub LabelPerfil_Click()
    Dim url As String
    url = "https://github.com/guerrahgv"
    
    On Error Resume Next
    Shell "cmd /c start " & url, vbHide
    On Error GoTo 0
End Sub


