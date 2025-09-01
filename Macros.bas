Attribute VB_Name = "Módulo1"
'MACROS QUE AYUDA A TRAER LA INFORMACION EN LA PAGINA DE MENU

Sub actualicacion_caja_borrador()

    'Dim cont As Long
    'Dim ultimalinea As Long
    Dim terminal As Variant '(valor de la tpv que se esta buscando)
    Dim caja As Variant
    Dim rango As Variant
    
    'cuenta la ultima linea de los datos en donde se pondra la info buscada(detalles)
    'ultimalinea = Sheets("MENU").Range("E3")
    
    
    
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
   
    'En cual linea empieza
    'For cont = 2 To ultimalinea
     
    ' donde se encuentra la terminal (key)(ejemplo=buscarv(terminal)
        terminal = Sheets("MENU").Range("E3")
        
    'Se aplica la función buscarv
        idterminal = Application.VLookup(terminal, rango, 2, False)
        
        
        
    'Si no hay info (funcion si.error)
        If IsError(idterminal) Then
            'dterminal = "error"
        End If
        
     'Se define en que columna se pone el resultado
        Sheets("Menu").Range("E5") = idterminal
       
             
  

End Sub



Sub actualicacion_caja()

    'Dim cont As Long
    'Dim ultimalinea As Long
    Dim terminal As Variant '(valor de la tpv que se esta buscando)
    Dim caja As Variant
    Dim rango As Variant
    
    'cuenta la ultima linea de los datos en donde se pondra la info buscada(detalles)
    'ultimalinea = Sheets("MENU").Range("E3")
    
    
    
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
   
    'En cual linea empieza
    'For cont = 2 To ultimalinea
     
    ' donde se encuentra la terminal (key)(ejemplo=buscarv(terminal)
        terminal = Sheets("MENU").Range("E3")
        
    'Se aplica la función buscarv
        caja = Application.VLookup(terminal, rango, 2, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(caja) Then
            idterminal = "error"
        End If
        
     'Se define en que columna se pone el resultado
        Sheets("Menu").Range("E5") = caja
        'Next cont
      
        

End Sub



Sub actualicacion_modelo()

    'Dim cont As Long
    'Dim ultimalinea As Long
    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim modelo As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
    
    'cuenta la ultima linea de los datos en donde se pondra la info buscada(detalles)
    'ultimalinea = Sheets("MENU").Range("E3")
    
    
    
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
   
    'En cual linea empieza
    'For cont = 2 To ultimalinea
     
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
        
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        modelo = Application.VLookup(terminal, rango, 3, False)
        
        
        
    'Si no hay info (funcion si.error)
        If IsError(modelo) Then
            modelo = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("E7") = modelo
       
  

End Sub


Sub actualicacion_ubistock()

    
    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim ubistock As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
        
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        ubistock = Application.VLookup(terminal, rango, 4, False)
        
        
        
    'Si no hay info (funcion si.error)
        If IsError(ubistock) Then
            ubistock = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("H3") = ubistock
         

End Sub


Sub actualicacion_fechentrada()


    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim fechentrada As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          'Se efinen las variables
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
                
        'Se aplica la funcion y se asigna a la variable buscada
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        fechentrada = Application.VLookup(terminal, rango, 5, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(fechentrada) Then
            fechentrada = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("H5") = fechentrada
        Range("H5").NumberFormat = "dd mmmm yyyy"
             
End Sub

    
Sub actualicacion_estatus()


    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim estatus As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          'Se efinen las variables
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
                
        'Se aplica la funcion y se asigna a la variable buscada
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        estatus = Application.VLookup(terminal, rango, 6, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(estatus) Then
            estatus = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("H7") = estatus
    
             
End Sub


Sub actualicacion_entrega()


    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim entrega As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          'Se efinen las variables
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
                
        'Se aplica la funcion y se asigna a la variable buscada
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        entrega = Application.VLookup(terminal, rango, 7, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(entrega) Then
            entrega = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("k3") = entrega
      
             
End Sub

  
Sub actualicacion_ubioperacion()


    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim entrega As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          'Se efinen las variables
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
                
        'Se aplica la funcion y se asigna a la variable buscada
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        ubioperacion = Application.VLookup(terminal, rango, 8, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(ubioperacion) Then
            ubioperacion = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("k5") = ubioperacion
     
             
End Sub

  
Sub actualicacion_fechasalida()


    Dim terminal As Variant '(valor de la tpv que se esta buscando, con respecto a la llave)
    Dim fechasalida As Variant '(La llave de la terminal buscada que se esta buscando)
    Dim rango As Variant      '(rango de la busqueda)
          
          'Se efinen las variables
    'Se genera el rango a donde se ira a extraer la info
    
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
        
    ' donde se encuentra la terminal (key)
        terminal = Sheets("MENU").Range("E3")
                
        'Se aplica la funcion y se asigna a la variable buscada
    'Se aplica la función buscarv ' este codigo se corrige de acuerdo con las columna de la cual se extrae la info de la hoja pagos
        fechasalida = Application.VLookup(terminal, rango, 9, False)
        
        
    'Si no hay info (funcion si.error)
        If IsError(fechasalida) Then
            fechasalida = "error"
        End If
        
     'Se define en que columna se pone el resultado 'esta columna siempre se cambia por que se imprime resultado
        Sheets("Menu").Range("k7") = fechasalida
      
             
End Sub



Sub BUSCAR()

    Call actualicacion_caja
    Call actualicacion_modelo
    Call actualicacion_ubistock
    Call actualicacion_fechentrada
    Call actualicacion_estatus
    Call actualicacion_entrega
    Call actualicacion_ubioperacion
    Call actualicacion_fechasalida

End Sub




'MACROS QUE AYUDA A ACTUALIZAR LA INFOMACION CUANDO SE DA UNA TPV A UN COMERCIAL

'Sub Macro4()
    ' Declarar variables
    'Dim valor_buscado As Range
    'Dim valor As Variant

    'Sheets("MENU").Range("E3").Select
   'Application.CutCopyMode = False
   'Selection.Copy
   'Sheets("INVENTARIO").Select

    ' Copiar el valor de la celda E3
    'Sheets("MENU").Range("E3").Copy

    ' Buscar el valor en la hoja INVENTARIO
    'valor = Sheets("MENU").Range("H7").Value
    
       
    'Set valor_buscado = Sheets("INVENTARIO").Cells.Find(What:=valor, After:=ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

    ' Verificar si se encontró el valor
    'If Not valor_buscado Is Nothing Then
        ' Seleccionar la celda encontrada
        'valor_buscado.Select
        'ActiveWindow.SmallScroll ToRight:=7
        ' Pegar valores
        'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    'Else
        'MsgBox "El valor no se encontró en la hoja INVENTARIO."
    'End If
    
    
Sub ActEstatus()
'
    ' Paso 1: Ir a la hoja MENU
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de E3 llamado TERMINAL
    Dim terminal As Variant
    terminal = Range("E3").Value
    
    ' Paso 3: Ir a la hoja INVENTARIO y buscar el valor de la celda E3
    Sheets("INVENTARIO").Activate
    Dim valor_buscado As Range
    Set valor_buscado = Cells.Find(What:=terminal, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not valor_buscado Is Nothing Then
        ' Paso 4: Desplazarse 5 celdas a la derecha
        valor_buscado.Offset(, 5).Select
        ' Paso 5: Regresar a la hoja MENU
        Sheets("MENU").Activate
        
        ' Paso 6: Copiar la celda H7 y pegar en INVENTARIO
        Range("H7").Copy
        Sheets("INVENTARIO").Activate
        valor_buscado.Offset(, 5).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgBox "El valor no se encontró en la hoja INVENTARIO."
    End If
    
End Sub


Sub ActEntrega()
'
    ' Paso 1: Ir a la hoja MENU
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de E3 llamado TERMINAL
    Dim terminal As Variant
    terminal = Range("E3").Value
    
    ' Paso 3: Ir a la hoja INVENTARIO y buscar el valor de la celda E3
    Sheets("INVENTARIO").Activate
    Dim valor_buscado As Range
    Set valor_buscado = Cells.Find(What:=terminal, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not valor_buscado Is Nothing Then
        ' Paso 4: Desplazarse 6 celdas a la derecha
        valor_buscado.Offset(, 6).Select
        ' Paso 5: Regresar a la hoja MENU
        Sheets("MENU").Activate
        
        ' Paso 6: Copiar la celda k3 y pegar en INVENTARIO
        Range("K3").Copy
        Sheets("INVENTARIO").Activate
        valor_buscado.Offset(, 6).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgBox "El valor no se encontró en la hoja INVENTARIO."
    End If
    
End Sub





Sub ActUbicacion()
'
    ' Paso 1: Ir a la hoja MENU
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de E3 llamado TERMINAL
    Dim terminal As Variant
    terminal = Range("E3").Value
    
    ' Paso 3: Ir a la hoja INVENTARIO y buscar el valor de la celda E3
    Sheets("INVENTARIO").Activate
    Dim valor_buscado As Range
    Set valor_buscado = Cells.Find(What:=terminal, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not valor_buscado Is Nothing Then
        ' Paso 4: Desplazarse 7 celdas a la derecha
        valor_buscado.Offset(, 7).Select
        ' Paso 5: Regresar a la hoja MENU
        Sheets("MENU").Activate
        
        ' Paso 6: Copiar la celda k5 y pegar en INVENTARIO
        Range("K5").Copy
        Sheets("INVENTARIO").Activate
        valor_buscado.Offset(, 7).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgBox "El valor no se encontró en la hoja INVENTARIO."
    End If
    
End Sub


Sub ActFechSalida()
'
    ' Paso 1: Ir a la hoja MENU
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de E3 llamado TERMINAL
    Dim terminal As Variant
    terminal = Range("E3").Value
    
    ' Paso 3: Ir a la hoja INVENTARIO y buscar el valor de la celda E3
    Sheets("INVENTARIO").Activate
    Dim valor_buscado As Range
    Set valor_buscado = Cells.Find(What:=terminal, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not valor_buscado Is Nothing Then
        ' Paso 4: Desplazarse 8 celdas a la derecha
        valor_buscado.Offset(, 8).Select
        ' Paso 5: Regresar a la hoja MENU
        Sheets("MENU").Activate
        
        ' Paso 6: Copiar la celda k6 y pegar en INVENTARIO
        Range("K7").Copy
        Sheets("INVENTARIO").Activate
        valor_buscado.Offset(, 8).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgBox "El valor no se encontró en la hoja INVENTARIO."
    End If
    
End Sub


Sub ActInventario()

    Call ActFechSalida
    Call ActUbicacion
    Call ActUbicacion
    Call ActEntrega
    Call ActEstatus
    Call ActFechSalida
    
End Sub




'MACROS QUE AYUDA A BUSCAR LA INFORMACIÓN EN LA HOJA DE PAGOS. SI NO ESTA SE AGREGA Y SI ESTA SE ACTUALIZA




Sub PgTerminal()

    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("E15").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "B").End(xlUp).Row + 1
    Range("B" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
    
End Sub



Sub PgNomina()

    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("E17").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "C").End(xlUp).Row
    Range("C" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
    
End Sub


Sub PgArea()

    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("E19").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "D").End(xlUp).Row
    Range("D" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
    
End Sub




Sub PgComercial()

    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("H15").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "E").End(xlUp).Row
    Range("E" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
    
End Sub



Sub PgEstPago()


    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("H17").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "F").End(xlUp).Row
    Range("F" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste

End Sub
    
Sub PgFecPago()



    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("H19").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "G").End(xlUp).Row
    Range("G" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
End Sub



Sub PgCosto()


    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("k15").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "H").End(xlUp).Row
    Range("H" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste
    
End Sub


Sub PgFactura()



    Dim valor As Variant

    ' Paso 1: Ir a la hoja "MENU"
    Sheets("MENU").Activate
    
    ' Paso 2: Copiar el valor de la celda E15
    valor = Range("k17").Copy
    
    ' Paso 3: Ir a la hoja "pagos"
    Sheets("PAGOS").Activate
    
    ' Paso 4: Ir a la última fila de la columna B
    Dim ultimaFila As Long
    ultimaFila = Cells(Rows.Count, "I").End(xlUp).Row
    Range("I" & ultimaFila).Select
    
    ' Paso 5: Pegar el valor
    ActiveSheet.Paste

End Sub

Sub PgActualizar()

    Call PgTerminal
    Call PgNomina
    Call PgArea
    Call PgComercial
    Call PgEstPago
    Call PgFecPago
    Call PgCosto
    Call PgFactura
    
End Sub


Sub STOCK()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim STOCK As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
        
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
    
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        STOCK = Application.VLookup(terminal, rango, 4, False)
    
        If IsError(STOCK) Then
            STOCK = "-"
        End If
        
     
        Sheets("DETALLES").Cells(cont, 3) = STOCK
        Next cont
     
'
End Sub



Sub ESTATUSTERMINAL()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim estatusterm As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
    
    For cont = 2 To ultimalinea
            terminal = Sheets("DETALLES").Cells(cont, 1)
        estatusterm = Application.VLookup(terminal, rango, 6, False)
        
        If IsError(estatustermin) Then
            estatusterm = "-"
        End If
        
           Sheets("DETALLES").Cells(cont, 4) = estatusterm
        Next cont
     
'
End Sub



Sub UBICACIONOPERACION()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim uboperacion As Variant
    Dim terminal As Variant
    Dim rango As Variant
   
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
    
    
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        ubioperacion = Application.VLookup(terminal, rango, 8, False)
        
   
        If IsError(ubioperacion) Then
            ubioperacion = "-"
        End If
        

        Sheets("DETALLES").Cells(cont, 5) = ubioperacion
        Next cont
     
'
End Sub




Sub fechasalida()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim fchsalida As Variant
    Dim terminal As Variant
    Dim rango As Variant
  
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("INVENTARIO").Range("B2:J1990")
    
            
    For cont = 2 To ultimalinea
           terminal = Sheets("DETALLES").Cells(cont, 1)
           fchsalida = Application.VLookup(terminal, rango, 9, False)
        
   
        If IsError(fchsalida) Then
            fchsalida = "-"
        End If
        
        Sheets("DETALLES").Cells(cont, 6) = fchsalida
        Next cont
     
'
End Sub



Sub AREA()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim AREA As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")
   
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        AREA = Application.VLookup(terminal, rango, 3, False)
    
        If IsError(AREA) Then
            AREA = "-"
        End If
    
        Sheets("DETALLES").Cells(cont, 7) = AREA
        Next cont
     
'
End Sub




Sub NOMINACOMERCIAL()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim nomcomercial As Variant
    Dim terminal As Variant
    Dim rango As Variant
   
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")

    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        nomcomercial = Application.VLookup(terminal, rango, 2, False)
  
        If IsError(nomcomercial) Then
            nomcomercial = "-"
        End If
     
        Sheets("DETALLES").Cells(cont, 8) = nomcomercial
        Next cont
     
'
End Sub



Sub NOMBRECOMERCIAL()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim nombre As Variant
    Dim terminal As Variant
    Dim rango As Variant
   
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")
    
    
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
    
        nombre = Application.VLookup(terminal, rango, 4, False)
     
        If IsError(nombre) Then
            nombre = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 9) = nombre
        Next cont
     
'
End Sub


Sub ESTATUSPAGOS()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim estatuspago As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")
    
    
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        
        estatuspago = Application.VLookup(terminal, rango, 5, False)
        
    
        If IsError(estatuspago) Then
            estatuspago = "-"
        End If
     
        Sheets("DETALLES").Cells(cont, 10) = estatuspago
        Next cont
     
'
End Sub


Sub FECHAPAGO()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim fchpago As Variant
    Dim terminal As Variant
    Dim rango As Variant
   
   ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")
    
    
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
        
        fchpago = Application.VLookup(terminal, rango, 6, False)
        
    
        If IsError(fchpago) Then
            fchpago = "-"
        End If
     
        Sheets("DETALLES").Cells(cont, 11) = fchpago
        Next cont
     
'
End Sub


Sub COSTOTERM()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim costo As Variant
    Dim terminal As Variant
    Dim rango As Variant
   
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("PAGOS").Range("B2:I1990")
    
   
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
    
        costo = Application.VLookup(terminal, rango, 7, False)
     
        If IsError(costo) Then
            costo = "-"
        End If
    
        Sheets("DETALLES").Cells(cont, 12) = costo
        Next cont
     
'
End Sub

Sub DIMCLIENTES()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim cliente As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("CLIENTES").Range("B2:D1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        cliente = Application.VLookup(terminal, rango, 3, False)
    
        If IsError(cliente) Then
            cliente = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 13) = cliente
        Next cont
     
'
End Sub



Sub NOMBREUSUARIO()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim usuario As Variant
    Dim terminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        usuario = Application.VLookup(terminal, rango, 4, False)
    
        If IsError(usuario) Then
            usuario = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 14) = usuario
        Next cont
     
'
End Sub



Sub CLABETOKAPAY()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim usuario As Variant
    Dim clabe As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        clabe = Application.VLookup(terminal, rango, 2, False)
    
        If IsError(clabe) Then
            clabe = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 15) = clabe
        Next cont
     
'
End Sub



'Sub CLABEBANCO() se quitara porque no tiene información esa columna
'

    'Dim cont As Long
    'Dim ultimalinea As Long
    'Dim usuario As Variant
    'Dim clabe As Variant
    'Dim rango As Variant
    
    'ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    'Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    'For cont = 2 To ultimalinea
        'terminal = Sheets("DETALLES").Cells(cont, 1)
   
        'clabe = Application.VLookup(terminal, rango, 3, False)
    
        'If IsError(clabe) Then
            'clabe = "-"
        'End If
      
        'Sheets("DETALLES").Cells(cont, 16) = clabe
        'Next cont
        
        
'End Sub


Sub ESTATUSUSUARIO()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim usuario As Variant
    Dim estusuario As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        estusuario = Application.VLookup(terminal, rango, 5, False)
    
        If IsError(estusuario) Then
            estusuario = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 17) = estusuario
        Next cont
     
'
End Sub



Sub ESTATUSEXPEDIENTE()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim usuario As Variant
    Dim estexpediente As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        estexpediente = Application.VLookup(terminal, rango, 6, False)
    
        If IsError(estexpediente) Then
            estexpediente = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 18) = estexpediente
        Next cont
     
'
End Sub







Sub ESTATUSTERMINAL1()
'

    Dim cont As Long
    Dim ultimalinea As Long
    Dim usuario As Variant
    Dim estterminal As Variant
    Dim rango As Variant
    
    ultimalinea = Sheets("DETALLES").Range("A" & Rows.Count).End(xlUp).Row
    Set rango = Sheets("USUARIOS").Range("B2:H1990")
           
    For cont = 2 To ultimalinea
        terminal = Sheets("DETALLES").Cells(cont, 1)
   
        estterminal = Application.VLookup(terminal, rango, 7, False)
    
        If IsError(estterminal) Then
            estterminal = "-"
        End If
      
        Sheets("DETALLES").Cells(cont, 19) = estterminal
        Next cont
     
'
End Sub



Sub ACTUALIZACIONDETALLES()

    Call STOCK
    Call ESTATUSTERMINAL
    Call UBICACIONOPERACION
    Call fechasalida
    Call AREA
    Call NOMINACOMERCIAL
    Call NOMBRECOMERCIAL
    Call ESTATUSPAGOS
    Call FECHAPAGO
    Call COSTOTERM
    Call DIMCLIENTES
    Call NOMBREUSUARIO
    Call CLABETOKAPAY
    Call ESTATUSUSUARIO
    Call ESTATUSEXPEDIENTE
    Call ESTATUSTERMINAL1
    
End Sub

