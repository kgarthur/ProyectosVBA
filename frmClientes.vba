Dim ArchivoIMG As String

Private Sub btnAgregar_Click()
Dim i As Integer

If cboNombre.Text = "" Then
MsgBox "Nombre inválido", vbInformation + vbOKOnly
cboNombre.SetFocus
Exit Sub
End If

If Not (Mid(cboNombre.Text, 1, 1) Like "[a-z]" Or Mid(cboNombre.Text, 1, 1) Like "[A-Z]") Then
MsgBox "Nombre inválido", vbInformation + vbOKOnly
cboNombre.SetFocus
Exit Sub
End If

For i = 2 To Len(cboNombre.Text)
If Mid(cboNombre.Text, i, 1) Like "#" Then
MsgBox "Nombre inválido", vbInformation + vbOKOnly
cboNombre.SetFocus
Exit Sub
End If
Next

Sheets("Clientes").Activate

Dim fCliente As Integer
fCliente = nCliente(cboNombre.Text)

If fCliente = 0 Then
Do While Not IsEmpty(ActiveCell)
ActiveCell.Offset(1, 0).Activate ' si el registro no existe, se va al final.
Loop
Else
Cells(fCliente, 1).Select ' cuando ya existe el registro, cumple esta condición.
End If


'Aqui es cuando agregamos o modificamos el registro
Application.ScreenUpdating = False
ActiveCell = cboNombre
ActiveCell.Offset(0, 1) = txtDireccion
ActiveCell.Offset(0, 2) = txtTelefono
ActiveCell.Offset(0, 3) = txtID
ActiveCell.Offset(0, 4) = txtEmail
ActiveCell.Offset(0, 5) = ArchivoIMG



Application.ScreenUpdating = True


LimpiarFormulario

cboNombre.SetFocus

End Sub
Private Sub btnEliminar_Click()
Dim fCliente As Integer
fCliente = nCliente(cboNombre.Text)

If fCliente = 0 Then
MsgBox "El cliente que usted quiere eliminar no existe", vbInformation + vbOKOnly
cboNombre.SetFocus
Exit Sub
End If

If MsgBox("¿Seguro que quiere eliminar este cliente?", vbQuestion + vbYesNo) = vbYes Then

Cells(fCliente, 1).Select

ActiveCell.EntireRow.Delete

LimpiarFormulario

MsgBox "Cliente eliminado", vbInformation + vbOKOnly
cboNombre.SetFocus

End If

End Sub
Private Sub btnCerrar_Click()
End
End Sub


Private Sub cboNombre_Change()
On Error Resume Next


If nCliente(cboNombre.Text) <> 0 Then

Sheets("Clientes").Activate

Cells(cboNombre.ListIndex + 2, 1).Select
txtDireccion = ActiveCell.Offset(0, 1)
txtTelefono = ActiveCell.Offset(0, 2)
txtID = ActiveCell.Offset(0, 3)
txtEmail = ActiveCell.Offset(0, 4)

fotografia.Picture = LoadPicture("")
fotografia.Picture = LoadPicture(ActiveCell.Offset(0, 5))


ArchivoIMG = ActiveCell.Offset(0, 5)

Else
txtDireccion = ""
txtTelefono = ""
txtID = ""
txtEmail = ""
ArchivoIMG = ""
fotografia.Picture = LoadPicture("")
End If
End Sub
Private Sub cboNombre_Enter()
CargarLista
End Sub
Sub CargarLista()
cboNombre.Clear

Sheets("Clientes").Select
Range("A2").Select 'CELDA DONDE INICIA TU PRIMER REGISTRO
Do While Not IsEmpty(ActiveCell)
cboNombre.AddItem ActiveCell.Value
ActiveCell.Offset(1, 0).Select
Loop
End Sub

Sub LimpiarFormulario()
CargarLista

cboNombre = ""
txtDireccion = ""
txtTelefono = ""
txtID = ""
txtEmail = ""
ArchivoIMG = ""
End Sub
Private Sub btnImagen_Click()
On Error Resume Next

ArchivoIMG = Application.GetOpenFilename("Imágenes jpg,*.jpg,Imágenes bmp,*.bmp", 0, "Seleccionar Imágen para Reegistro de Clientes")
fotografia.Picture = LoadPicture("")
fotografia.Picture = LoadPicture(ArchivoIMG)

End Sub
