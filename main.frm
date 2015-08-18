VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "gjf2cmc"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btopen 
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox txtoutfile 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   480
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btEjecutar 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton btbrowse 
      Caption         =   "..."
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   612
   End
   Begin VB.TextBox txtinfile 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Output File:"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Input File:"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   852
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btopen_Click()
    If txtoutfile.Text <> "" Then
        Shell "notepad " & txtoutfile.Text, vbNormalFocus
    End If
End Sub

Private Sub btbrowse_Click()
    Dim n As Boolean
    
    n = abrirfichero()
End Sub

Function abrirfichero() As Boolean
    
    abrirfichero = True
            
    On Error GoTo cancela
    CDialog.Filter = "Archivos de Gaussian (*.gjf)|*.gjf|" & "Todos los archivos (*.*)|*.*"
    CDialog.FilterIndex = 1
    CDialog.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNNoChangeDir Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    CDialog.CancelError = True
    CDialog.InitDir = App.Path
    CDialog.ShowOpen
    txtinfile.Text = CDialog.FileName
    If txtinfile.Text = "" Then
        abrirfichero = False
        Exit Function
    End If
    
    txtoutfile.Text = getPath(txtinfile.Text) & getBasename(txtinfile.Text, False) & ".dat"
    btopen.Enabled = False
    
    Exit Function
cancela:
    If (Err.Number = cdlCancel) Then
        Mensaje = "Seleccion cancelada"
        Estilo = vbInformation + vbDefaultButton1
    Else
        Mensaje = "Se ha producido un error al abrir el archivo"
        Estilo = vbCritical + vbDefaultButton1
    End If
    Titulo = "Selección del fichero de gaussian (*.gjf)"
    respuesta = MsgBox(Mensaje, Estilo, Titulo)
    abrirfichero = False
End Function

'*******************************************************************
'    FUNCIONES PARA CONVERTIR LOS DATOS DE UN FICHERO GJF A LOS DEL UCA-CMC
'*******************************************************************
Private Sub btEjecutar_Click()
    Dim fichero As String
    Dim fila As Integer, i As Integer, j As Integer, k As Integer
    Dim cadin As String, ext As String
    Dim splitbloque() As String, splitgeom() As String, cm() As String, cm1() As String, geom() As String, geom1() As String, geometry() As String
    Dim splitbloquefbase() As String, splitbase() As String, base() As String, base1() As String
    Dim fb() As fbase, nfb As Integer
    Dim natom As Integer, carga As Integer, multiplicidad As Integer, numelect As Integer, noca As Integer, nocb As Integer, factor As Integer
    Dim fichout As String
    Dim fout As Integer
    Dim formula As String
            
    Dim numcentro As Integer, partangular As String, numsfb As Integer
            
    ext = getExt(txtinfile.Text)
    If LCase(ext) = "gjf" Then
        fichero = txtinfile.Text
        
        cadin = leetodo(fichero, False)
        
        splitbloque = Split(cadin, vbLf & vbCr, -1, vbTextCompare)
        
        splitgeom = Split(splitbloque(2), vbCr, -1, vbTextCompare)
        
        natom = CInt(UBound(splitgeom) - 1)
        cm = Split(splitgeom(0), " ", -1, vbTextCompare)
        Call limpiaBlancos(cm, cm1)
        carga = CInt(cm1(0))
        multiplicidad = CInt(cm1(1))
        
        'geometría
        ReDim geometry(1 To 7, 1 To natom)
        For j = 1 To natom
           For k = 1 To 7
               geometry(k, j) = ""
           Next k
        Next j
        For j = 1 To natom
           geom = Split(splitgeom(j), " ", -1, vbTextCompare)
           Call limpiaBlancos(geom, geom1)
           On Error Resume Next
           For k = 1 To 7
                geometry(k, j) = geom1(k - 1)
           Next k
        Next j
        
        'funciones de base
        If splitbloque(3) <> "" Then
            splitbloquefbase = Split(splitbloque(3), "****" & vbCr, -1, vbTextCompare)
            nfb = 0
            For j = 0 To UBound(splitbloquefbase)
                splitbase = Split(splitbloquefbase(j), vbCr, -1, vbTextCompare)
                base = Split(splitbase(0), " ", -1, vbTextCompare)
                Call limpiaBlancos(base, base1)
                
                If IsNumeric(base1(0)) Then  'se le aplica al atomo con este numero
                    numcentro = CInt(base1(0))
                    Call crea_base(splitbase, numcentro, nfb, fb)
                Else 'se le aplica a todos los átomos con ese símbolo ¡¡¡¡No está hecho!!!!
                    For k = 1 To natom
                        If base1(0) = geometry(1, k) Then
                            numcentro = k
                            Call crea_base(splitbase, numcentro, nfb, fb)
                        End If
                    Next k
                End If
                
            Next j
        Else
            MsgBox "Debes leer un fichero gjf donde se ha introducido las bases dadas por la opción gfinput del Gaussian09", vbCritical
            Exit Sub
        End If
        
        'ocupación
        numelect = 0
        For j = 1 To natom
            numelect = numelect + getZnuc(geometry(1, j))
        Next j
        numelect = numelect - carga
        If isPar(numelect) Then 'todos están aparecados
            noca = numelect / 2
            nocb = noca
            factor = multiplicidad - 1
        Else 'este es el caso de los desapareados de un sólo electrón
            noca = (numelect - 1) / 2
            nocb = noca
            noca = noca + 1
            factor = multiplicidad - 2
        End If
        If multiplicidad > 2 Then 'en caso de una mayor multiplicidad quitamos "factor/2" de los beta y los ponemos en los alfa
            noca = noca + factor / 2
            nocb = nocb - factor / 2
        End If
        
        'escritura del fichero para el uca-cmc
        'fichout = getPath(fichero) & getBasename(fichero, False) & ".dat"
        fichout = txtoutfile
        fout = FreeFile
        Open (fichout) For Output As #fout
            Print #fout, "[NUCLEOS]"
            Print #fout, "Nucleos=" & natom
            For j = 1 To natom
                Print #fout, "Atomo(" & j & ",centro)=" & j
                Print #fout, "Atomo(" & j & ",znuc)=" & CStr(getZnuc(geometry(1, j)))
            Next j
            Print #fout, ""
            Print #fout, "[FBASE]"
            Print #fout, "numero de funciones de base:"
            Print #fout, "NumFB=" & nfb
            Print #fout, "tipo de las funciones de base:"
            Print #fout, "TipoFB=1"
            Print #fout, "Normalizacion de las SUBfunciones de base, =0 coef. para SFB no normalizadas, =1 para SFB normalizadas"
            Print #fout, "NORMSFB=0"
            Print #fout, "Empleo de funciones recortadas (TSTOs),  =1 son TSTO, =0 son STO, BO o SBO"
            Print #fout, "Recortadas=1"
            Print #fout, "Cálculo de integrales nulas (=0 se calculan, =1 se saltan)"
            Print #fout, "Bielec=0"
            Print #fout, ""
            For j = 1 To nfb
                Print #fout, "FBase(" & j & ",Centro)=" & fb(j).centro
                Print #fout, "FBase(" & j & ",ParteAngular)=" & fb(j).parteangular
                Print #fout, "FBase(" & j & ",numgauss)=" & fb(j).numgss
                formula = ""
                For k = 1 To fb(j).numgss
                    Print #fout, "FBase(" & j & "," & k & ",Coeficiente)=" & fb(j).subgss(k).coef
                    Print #fout, "FBase(" & j & "," & k & ",Parametro1)=" & fb(j).subgss(k).p1
                    Print #fout, "FBase(" & j & "," & k & ",Parametro2)=" & fb(j).subgss(k).p2
                    Print #fout, "FBase(" & j & "," & k & ",Parametro3)=" & fb(j).subgss(k).p3
                    If k = 1 Then
                        Select Case fb(j).subgss(k).p1
                            Case 0: formula = fb(j).subgss(k).coef & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                            Case 1: formula = fb(j).subgss(k).coef & "r" & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                            Case Default: formula = fb(j).subgss(k).coef & "r^" & fb(j).subgss(k).p1 & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                        End Select
                    Else
                        Select Case fb(j).subgss(k).p1
                            Case 0: formula = formula & " + " & fb(j).subgss(k).coef & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                            Case 1: formula = formula & " + " & fb(j).subgss(k).coef & "r" & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                            Case Default: formula = formula & " + " & fb(j).subgss(k).coef & "r^" & fb(j).subgss(k).p1 & " Exp(-" & fb(j).subgss(k).p3 & " r^2)"
                        End Select
                    End If
                Next k
                Print #fout, "FBase(" & j & ",RadioCorte0)=0"
                Print #fout, "FBase(" & j & ",AnchuraCorte0)=0"
                Print #fout, "FBase(" & j & ",RadioCorte1)=0"
                Print #fout, "FBase(" & j & ",AnchuraCorte1)=0"
                Print #fout, "FBase(" & j & ",TipoIntCorte)=0"
                Print #fout, "FBase(" & j & ",Formula)=" & formula
                Print #fout, ""
            Next j
            
            Print #fout, "[CENTROS]"
            Print #fout, "NumCentros=" & natom
            For j = 1 To natom
                Print #fout, "Centro(" & j & ",x)=" & CStr(geometry(2, j))
                Print #fout, "Centro(" & j & ",y)=" & CStr(geometry(3, j))
                Print #fout, "Centro(" & j & ",z)=" & CStr(geometry(4, j))
            Next j
            Print #fout, ""
        
            Print #fout, "[METODOINTEGRACION]"
            Print #fout, "Metodo=GSS"
            Print #fout, "Semilla=1"
            Print #fout, "NumPuntos=1000000"
            Print #fout, "Fichgss="
            Print #fout, ""
            
            Print #fout, "[INTEGRALES]"
            Print #fout, "FicheroInt =" & getBasename(fichero, False)
            Print #fout, "Integrales = 0"
            Print #fout, "TipoInt = j"
            Print #fout, "Indice1 = 1"
            Print #fout, "Indice2 = 1"
            Print #fout, "Indice3 = 1"
            Print #fout, "Indice4 = 1"
            Print #fout, "FPeso(NPolig) = 0"
            Print #fout, "FPeso(Inicio) = 0"
            Print #fout, "FPeso(fin) = 0"
            Print #fout, "FPeso(Incremento) = 0"
            Print #fout, ""
            
            Print #fout, "[METODOCALCULO]"
            Print #fout, "metodo = HF"
            Print #fout, "NmaxIter = 1000"
            Print #fout, "ControlConvergencia = 1"
            Print #fout, "Tolerancia = 0.0000001"
            Print #fout, "FactorAmortiguacion = 0.5"
            Print #fout, "IteracionComienzoAmortiguacion = 0"
            Print #fout, ""
            
            Print #fout, "[ESTADOELECTRONICO]"
            For j = 1 To noca 'numeros de ocupacion alfa siempre mayor o igual a nocb
                Print #fout, "NumOcupAlfa(" & j & ")=" & "1"
                If j <= nocb Then
                    Print #fout, "NumOcupBeta(" & j & ")=" & "1"
                Else
                    Print #fout, "NumOcupBeta(" & j & ")=" & "0"
                End If
            Next j
            For j = 1 To noca
                For k = 1 To noca
                    Print #fout, "CoefOcupAlfa(" & j & ", " & k & ") =0"
                    Print #fout, "CoefOcupBeta(" & j & ", " & k & ") =0"
                    Print #fout, "CoefOcupAlfa(" & j & ", " & k & ") =0"
                    Print #fout, "CoefOcupBeta(" & j & ", " & k & ") =0"
                Next k
            Next j
            Print #fout, ""
            
            Print #fout, "[POTENCIALEXT]"
            Print #fout, "Potencial(X) = 0"
            Print #fout, "Potencial(Y) = 0"
            Print #fout, "Potencial(z) = 0"
            Print #fout, ""
            
            Print #fout, "[PROPIEDADES]"
            Print #fout, "MomentoDipolar = 0"
            Print #fout, "Poblaciones = 0"
            Print #fout, ""
            
            Print #fout, "[RESULTADOS]"
            Print #fout, "Escritura(0) = 1"
            Print #fout, "Escritura(1) = 1"
            Print #fout, "Escritura(2) = 1"
            Print #fout, "Escritura(3) = 1"
            Print #fout, "Escritura(4) = 1"
            Print #fout, "Escritura(5) = 1"
            Print #fout, "Escritura(6) = 1"
            Print #fout, "Escritura(7) = 1"
            Print #fout, "Escritura(8) = 1"
            Print #fout, "Escritura(9) = 1"
            Print #fout, "Escritura(10) = 1"
            Print #fout, "Escritura(11) = 1"
            Print #fout, "Escritura(12) = 1"
            Print #fout, "Escritura(13) = 1"
        Close fout
        btopen.Enabled = True
     Else
        MsgBox "Debes leer un fichero gjf donde se ha introducido las bases dadas por la opción gfinput del Gaussian09", vbCritical
     End If

End Sub

Private Sub crea_base(splitbase() As String, numcentro As Integer, nfb As Integer, fb() As fbase)
    Dim i As Integer, line As Integer
    Dim base() As String, base1() As String
    Dim partangular As String
    Dim numgauss As Integer
    
    
    line = 0
    Do While line < UBound(splitbase) - 1
        line = line + 1
        base = Split(splitbase(line), " ", -1, vbTextCompare)
        Call limpiaBlancos(base, base1)
        partangular = base1(0)
        numgauss = CInt(base1(1))
        
        Select Case partangular
            Case "S"
                nfb = nfb + 1
                ReDim Preserve fb(1 To nfb)
                ReDim Preserve fb(nfb).subgss(1 To numgauss)
                fb(nfb).centro = numcentro
                fb(nfb).parteangular = 0
                fb(nfb).numgss = numgauss
                For k = 1 To numgauss
                    line = line + 1
                    base = Split(splitbase(line), " ", -1, vbTextCompare)
                    Call limpiaBlancos(base, base1)
                    fb(nfb).subgss(k).coef = Val(base1(1))
                    fb(nfb).subgss(k).p1 = 0  'se le pone siempre cero porque el cmc no lo utiliza con funciones gaussianas (son siempre gaussianas 1s)
                    fb(nfb).subgss(k).p2 = 0
                    fb(nfb).subgss(k).p3 = Val(base1(0))
                Next k
            Case "P"
                For l = 1 To 3 'son 3 p
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(1))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case "D"
                For l = 1 To 5 'son 5 d
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l + 3
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(1))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case "SP"
                nfb = nfb + 1
                ReDim Preserve fb(1 To nfb)
                ReDim Preserve fb(nfb).subgss(1 To numgauss)
                fb(nfb).centro = numcentro
                fb(nfb).parteangular = 0
                fb(nfb).numgss = numgauss
                For k = 1 To numgauss
                    line = line + 1
                    base = Split(splitbase(line), " ", -1, vbTextCompare)
                    Call limpiaBlancos(base, base1)
                    fb(nfb).subgss(k).coef = Val(base1(1))
                    fb(nfb).subgss(k).p1 = 0
                    fb(nfb).subgss(k).p2 = 0
                    fb(nfb).subgss(k).p3 = Val(base1(0))
                Next k
                line = line - numgauss
                For l = 1 To 3 'son 3 p
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(2))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case "SPD"
                nfb = nfb + 1
                ReDim Preserve fb(1 To nfb)
                ReDim Preserve fb(nfb).subgss(1 To numgauss)
                fb(nfb).centro = numcentro
                fb(nfb).parteangular = 0
                fb(nfb).numgss = numgauss
                For k = 1 To numgauss
                    line = line + 1
                    base = Split(splitbase(line), " ", -1, vbTextCompare)
                    Call limpiaBlancos(base, base1)
                    fb(nfb).subgss(k).coef = Val(base1(1))
                    fb(nfb).subgss(k).p1 = 0
                    fb(nfb).subgss(k).p2 = 0
                    fb(nfb).subgss(k).p3 = Val(base1(0))
                Next k
                line = line - numgauss
                For l = 1 To 3 'son 3 p
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(2))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                For l = 1 To 5 'son 5 d
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l + 3
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(3))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case "F"
                For l = 1 To 7 'son 7 f
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l + 8
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(1))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case "G"
                For l = 1 To 9 'son 9 g
                    nfb = nfb + 1
                    ReDim Preserve fb(1 To nfb)
                    ReDim Preserve fb(nfb).subgss(1 To numgauss)
                    fb(nfb).centro = numcentro
                    fb(nfb).parteangular = l + 15
                    fb(nfb).numgss = numgauss
                    For k = 1 To numgauss
                        line = line + 1
                        base = Split(splitbase(line), " ", -1, vbTextCompare)
                        Call limpiaBlancos(base, base1)
                        fb(nfb).subgss(k).coef = Val(base1(1))
                        fb(nfb).subgss(k).p1 = 0
                        fb(nfb).subgss(k).p2 = 0
                        fb(nfb).subgss(k).p3 = Val(base1(0))
                    Next k
                    line = line - numgauss
                Next l
                line = line + numgauss
            Case Else
                MsgBox "No está contemplada la parte angular " & partangular, vbCritical
                Exit Sub
        End Select
    Loop

End Sub

'*******************************************************************
'    FUNCIONES VARIAS
'*******************************************************************

'lee todo el fichero y lo mete en un string. Si crlf=true quita los retornos de carro de todas las lineas
Private Function leetodo(fichero As String, Optional crlf As Boolean) As String
    Dim cadin As String, f As Integer
    Dim tam As Long
    
    f = FreeFile
    Open fichero For Input As #f
    cadin = ""
    tam = LOF(f)
    cadin = Input$(tam, f)
    Close f
    
    If crlf = True Then
        leetodo = Replace(cadin, vbCrLf + " ", "", 1, -1, vbBinaryCompare)
    Else
        leetodo = cadin
    End If
End Function

'devuelve la cadena entre comp y fin. Si fin=NULL devuelve desde comp hasta el final. sensitive=true distingue entre may. y min.
Private Function getCad(cadin As String, comp As String, fin As String, sensitive As Boolean, Optional dir As Boolean) As String
    Dim pos As Long, poslen As Long, cadaux As String
    
    getCad = ""
    If cadin = "" Or comp = "" Then Exit Function
    If sensitive Then
        If dir Then
            pos = InStrRev(cadin, comp, -1, vbTextCompare) + Len(comp)
        Else
            pos = InStr(1, cadin, comp, vbTextCompare) + Len(comp)
        End If
    Else
        If dir Then
            pos = InStrRev(LCase(cadin), LCase(comp), -1, vbTextCompare) + Len(comp)
        Else
            pos = InStr(1, LCase(cadin), LCase(comp), vbTextCompare) + Len(comp)
        End If
    End If
    If pos = Len(comp) Then Exit Function  'no lo ha encontrado
    If fin = "" Then
        poslen = Len(cadin) - pos + 1
    Else
        cadaux = Mid(cadin, pos + 1, Len(cadin) - pos)
        If sensitive Then
            poslen = InStr(1, cadaux, fin, vbTextCompare)
         Else
            poslen = InStr(1, LCase(cadaux), LCase(fin), vbTextCompare)
        End If
    End If
    If pos > 0 And poslen > 0 Then
        getCad = Mid(cadin, pos, poslen)
    End If
End Function

'devuelve true si contine comp, sensitive=true distingue entre may. y min.
Private Function isCad(cadin As String, comp As String, sensitive As Boolean) As Boolean
    Dim pos As Long, poslen As Long, cadaux As String
    
    If cadin = "" Or comp = "" Then Exit Function
    If sensitive Then
        pos = InStr(1, cadin, comp, vbTextCompare)
    Else
        pos = InStr(1, LCase(cadin), LCase(comp), vbTextCompare)
    End If
    If pos Then
        isCad = True
    Else
        isCad = False
    End If
End Function

Private Function isPar(num As Integer) As Boolean
    isPar = True
    If (num Mod 2 <> 0) Then isPar = False
End Function

Private Function getNumber(cadin As String, comp As String, fin As String, sensitive As Boolean, Optional dir As Boolean) As Double
    Dim cadinn As String
    
    getNumber = 0
    cadinn = Trim(getCad(cadin, comp, fin, sensitive, dir))
    If IsNumeric(cadinn) Then
        getNumber = Val(cadinn)
    End If
 End Function
 
 'limpia los blancos en un array mono
Private Sub limpiaBlancos(arrin() As String, arrout() As String)
    Dim i As Integer, j As Integer
    
    j = -1
    For i = 0 To UBound(arrin)
        If arrin(i) <> "" And arrin(i) <> vbLf Then
            j = j + 1
            ReDim Preserve arrout(j)
            If Left(arrin(i), 1) = vbLf Then
                arrout(j) = Right(arrin(i), Len(arrin(i)) - 1)
            Else
                arrout(j) = arrin(i)
            End If
        End If
    Next i
End Sub

'nos dice las columnas que estan en blanco en la primera dimensión de un array bidimensional
Private Function columVacia(arrin() As String) As String
    Dim i As Integer, j As Integer, rows As Integer, cols As Integer, colmenos As Integer
    Dim cv As String
    Dim blanco As Boolean
    
    rows = UBound(arrin, 2)
    cols = UBound(arrin, 1)
    
    colmenos = 0
    For j = 0 To cols
        blanco = True
        For i = 0 To rows
            If arrin(j, i) <> "" Then
                blanco = False
                Exit For
            End If
        Next i
        If blanco = True Then
            colmenos = colmenos + 1
            cv = cv & " " & CStr(j)
        End If
    Next j
    columVacia = cv
End Function

Private Function IsColumVacia(col As String, cols As String) As Boolean
    Dim splitcols() As String
    Dim i As Integer
    
    IsColumVacia = False
    splitcols = Split(cols, " ", -1, vbTextCompare)
    For i = 0 To UBound(splitcols)
        If col = splitcols(i) Then
            IsColumVacia = True
            Exit For
        End If
    Next i
End Function
 
Private Sub linux2windows(fichero As String, original As Boolean)
    Dim f As Integer, g As Integer, cad As String, ficherobak As String, ficheroaux As String
    Dim longtotal As Long, longp As Long, bytesperstring As Long, ncads As Long, i As Long, pos As Long
    
    On Error GoTo error_Handler
    
    bytesperstring = 65000000 '65 Mb?
    longtotal = FileLen(fichero)
    If longtotal > bytesperstring Then
        longp = bytesperstring
        ncads = longtotal / bytesperstring
    Else
        longp = longtotal
        ncads = 1
    End If
        
    f = FreeFile
    Open fichero For Input As #f
    g = FreeFile
    ficherobak = getPath(fichero) & getBasename(fichero, False) & ".bak"
    Open ficherobak For Output As #g
    If ncads > 1 Then
        For i = 0 To ncads - 2
            cad = ""
            cad = Input$(longp, f)
            pos = InStr(1, cad, vbCrLf, vbBinaryCompare)
            If pos = 0 Then cad = Replace(cad, vbLf, vbCrLf, 1, -1, vbBinaryCompare)
            cad = Replace(cad, "\", "|", 1, -1, vbTextCompare)
            Print #g, cad
            DoEvents
        Next i
        cad = ""
        cad = Input$(longtotal - Seek(f), f)
        pos = InStr(1, cad, vbCrLf, vbBinaryCompare)
        If pos = 0 Then cad = Replace(cad, vbLf, vbCrLf, 1, -1, vbBinaryCompare)
        cad = Replace(cad, "\", "|", 1, -1, vbTextCompare)
        Print #g, cad
    Else
        cad = ""
        cad = Input$(longtotal, f)
        pos = InStr(1, cad, vbCrLf, vbBinaryCompare)
        If pos = 0 Then cad = Replace(cad, vbLf, vbCrLf, 1, -1, vbBinaryCompare)
        cad = Replace(cad, "\", "|", 1, -1, vbTextCompare)
        Print #g, cad
    End If
    
    Close #f
    Close #g
    
    If original = False Then
        Kill fichero
        Name ficherobak As fichero
    Else
        ficheroaux = getPath(fichero) & getBasename(fichero, False) & ".aux"
        Name fichero As ficheroaux
        Name ficherobak As fichero
        Name ficheroaux As ficherobak
    End If
        
    
    Exit Sub
    
error_Handler:
    MsgBox Err.Description, vbCritical
End Sub



Private Function getPath(File As String) As String
    Dim l As Long
    
    getPath = ""
    l = InStrRev(File, "\", -1, vbTextCompare)
    getPath = Left(File, l)
End Function

Private Function getBasename(File As String, ext As Boolean) As String
    Dim fileaux As String, l As Long
    
    getBasename = ""
    l = InStrRev(File, "\", -1, vbTextCompare)
    fileaux = Right(File, Len(File) - l)
    If fileaux <> "" Then
        If ext Then
            getBasename = fileaux
        Else
            l = InStrRev(fileaux, ".", -1, vbTextCompare)
            getBasename = Left(fileaux, l - 1)
        End If
    End If
End Function

Private Function getExt(File As String) As String
    Dim cad As String
    Dim pos As Integer
        
    cad = getBasename(File, True)
    pos = InStrRev(cad, ".", -1, vbTextCompare)
    If pos Then
        getExt = Mid(cad, pos + 1, Len(cad) - pos + 1)
    Else
        getExt = cad
    End If
End Function

Private Sub linux2windows1(fichero As String)
    Dim f As Integer, cad As String, pos As Long
    
    cad = leetodo(fichero)
    pos = InStr(1, cad, vbCrLf, vbBinaryCompare)
    If pos = 0 Then
        cad = Replace(cad, vbLf, vbCrLf, 1, -1, vbBinaryCompare)
    End If
    cad = Replace(cad, "\", "|", 1, -1, vbTextCompare)
    'cad = Replace(cad, "|" & vbCrLf & " |", "||" & vbCrLf & " ", 1, -1, vbBinaryCompare)
    
    f = FreeFile
    Open fichero For Output As #f
    Print #f, cad
    Close f
End Sub


Private Function ischar(c As String) As Boolean
    Dim n As Integer
    
    n = Asc(c)
    If n >= 48 And n <= 57 Then
        ischar = True
        Exit Function
    End If
    If n >= 65 And n <= 90 Then
        ischar = True
        Exit Function
    End If
    If n >= 97 And n <= 122 Then
        ischar = True
        Exit Function
    End If
    If n >= 128 And n <= 151 Then
        ischar = True
        Exit Function
    End If
    If n >= 153 And n <= 154 Then
        ischar = True
        Exit Function
    End If
    If n >= 160 And n <= 165 Then
        ischar = True
        Exit Function
    End If
End Function

'*******************************************************************
'    FUNCIONES PARA ATOMOS
'*******************************************************************
Private Function IsAtom(atom As String) As Boolean
    Const atomos As String = "H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr"
    If atom <> "" And InStr(1, atomos, atom, vbTextCompare) > 0 Then IsAtom = True
End Function

Public Function getSymbol(z As Integer) As String
    Dim atomo() As String
    Const atomos As String = "H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr"
    
    atomo = Split(atomos, " ", -1, vbTextCompare)
    getSymbol = atomo(z - 1)
End Function

Private Function getZnuc(atom As String) As Integer
    Dim atomo() As String, i As Integer
    Const atomos As String = "H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Sc Ti V Cr Mn Fe Co Ni Cu Zn Ga Ge As Se Br Kr Rb Sr Y Zr Nb Mo Tc Ru Rh Pd Ag Cd In Sn Sb Te I Xe Cs Ba La Ce Pr Nd Pm Sm Eu Gd Tb Dy Ho Er Tm Yb Lu Hf Ta W Re Os Ir Pt Au Hg Tl Pb Bi Po At Rn Fr Ra Ac Th Pa U Np Pu Am Cm Bk Cf Es Fm Md No Lr"
    
    atomo = Split(atomos, " ", -1, vbTextCompare)
    For i = 0 To UBound(atomo)
        If UCase(atomo(i)) = UCase(atom) Then
            getZnuc = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function getAngNum(ang As String) As Integer
    Dim pang() As String, i As Integer
    Const pangular As String = "s p d f g"
    
    pang = Split(pangular, " ", -1, vbTextCompare)
    For i = 0 To UBound(pang)
        If UCase(pang(i)) = UCase(ang) Then
            getAngNum = i
            Exit Function
        End If
    Next i
    
End Function

Public Sub parte_angular(dir As Boolean, num As Integer, pang As String)
    If dir = True Then
        Select Case num
            Case 0: pang = "S"
            Case 1: pang = "Px"
            Case 2: pang = "Py"
            Case 3: pang = "Pz"
            Case 4: pang = "D(x²-y²)"
            Case 5: pang = "Dxy"
            Case 6: pang = "Dxz"
            Case 7: pang = "Dyz"
            Case 8: pang = "D(3z²-r²)"
            Case 9: pang = "F(x³-3y²x)"
            Case 10: pang = "F(y³-3x²y)"
            Case 11: pang = "F(x²-y²)z"
            Case 12: pang = "Fxyz"
            Case 13: pang = "F(xr²-5xz²)"
            Case 14: pang = "F(yr²-5yz²)"
            Case 15: pang = "F(5z³-3r²z)"
        End Select
    Else
        Select Case pang
            Case "S": num = 0
            Case "Px": num = 1
            Case "Py": num = 2
            Case "Pz": num = 3
            Case "D(x²-y²)": num = 4
            Case "Dxy": num = 5
            Case "Dxz": num = 6
            Case "Dyz": num = 7
            Case "D(3z²-r²)": num = 8
            Case "F(x³-3y²x)": num = 9
            Case "F(y³-3x²y)": num = 10
            Case "F(x²-y²)z": num = 11
            Case "Fxyz": num = 12
            Case "F(xr²-5xz²)": num = 13
            Case "F(yr²-5yz²)": num = 14
            Case "F(5z³-3r²z)": num = 15
        End Select
    End If
End Sub

Public Sub conf_atomo_neutro(z As Integer, nu() As Byte)
    Dim i As Integer
    'ReDim nu(1 To 18)
    
    For i = 1 To 18
        nu(i) = 0
    Next i
      
    Select Case z
        Case 1 To 2
          nu(1) = z
        Case 3 To 4
          nu(1) = 2
          nu(2) = z - 2
        Case 5 To 10
          nu(1) = 2
          nu(2) = 2
          nu(3) = z - 4
        Case 11 To 12
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = z - 10
        Case 13 To 18
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = z - 12
        Case 19 To 20
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(7) = z - 18
        Case 21 To 30
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = z - 20
          nu(7) = 2
        Case 31 To 36
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = z - 30
        Case 37 To 38
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(11) = z - 36
        Case 39 To 48
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = z - 38
          nu(11) = 2
        Case 49 To 54
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(11) = 2
          nu(12) = z - 48
        Case 55 To 56
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(11) = 2
          nu(12) = 6
          nu(15) = z - 54
        Case 57 To 70
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = z - 56
          nu(11) = 2
          nu(12) = 6
          nu(15) = 2
        Case 71 To 80
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = 14
          nu(11) = 2
          nu(12) = 6
          nu(13) = z - 70
          nu(15) = 2
        Case 81 To 86
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = 14
          nu(11) = 2
          nu(12) = 6
          nu(13) = 10
          nu(15) = 2
          nu(16) = z - 80
        Case 87 To 88
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = 14
          nu(11) = 2
          nu(12) = 6
          nu(13) = 10
          nu(15) = 2
          nu(16) = 6
          nu(18) = z - 86
        Case 89 To 102
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = 14
          nu(11) = 2
          nu(12) = 6
          nu(13) = 10
          nu(14) = z - 88
          nu(15) = 2
          nu(16) = 6
          nu(18) = 2
        Case 103
          nu(1) = 2
          nu(2) = 2
          nu(3) = 6
          nu(4) = 2
          nu(5) = 6
          nu(6) = 10
          nu(7) = 2
          nu(8) = 6
          nu(9) = 10
          nu(10) = 14
          nu(11) = 2
          nu(12) = 6
          nu(13) = 10
          nu(14) = 14
          nu(15) = 2
          nu(16) = 6
          nu(17) = z - 102
          nu(18) = 2
    End Select
    
    rara = False 'configuraciones no estandar
    If z = 41 Or z = 42 Or z = 44 Or z = 45 Or z = 46 Or z = 47 _
       Or z = 57 Or z = 64 Or z = 78 Or z = 79 Or z = 89 Or z = 90 _
       Or z = 91 Or z = 92 Or z = 93 Or z = 96 Then
        Select Case z
            Case 41
                nu(9) = 4
                nu(11) = 1
            Case 42
                nu(9) = 5
                nu(11) = 1
            Case 44
                nu(9) = 7
                nu(11) = 1
            Case 45
                nu(9) = 8
                nu(11) = 1
            Case 46
                nu(9) = 10
                nu(11) = 0
            Case 47
                nu(9) = 10
                nu(11) = 1
            Case 57
                nu(10) = 0
                nu(13) = 1
            Case 64
                nu(10) = 7
                nu(13) = 1
            Case 78
                nu(13) = 9
                nu(15) = 1
            Case 79
                nu(13) = 10
                nu(15) = 1
            Case 89
                nu(14) = 0
                nu(17) = 1
            Case 90
                nu(14) = 0
                nu(17) = 2
            Case 91
                nu(14) = 2
                nu(17) = 1
            Case 92
                nu(14) = 3
                nu(17) = 1
            Case 93
                nu(14) = 4
                nu(17) = 1
            Case 96
                nu(14) = 7
                nu(17) = 1
        End Select
        rara = True
    End If
End Sub

