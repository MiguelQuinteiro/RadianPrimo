Attribute VB_Name = "Module1"

'PARA PODER GUARDAR IMAGEN DEL FORMULARIO
Public Declare Sub keybd_event _
                    Lib "user32" ( _
                        ByVal bVk As Byte, _
                        ByVal bScan As Byte, _
                        ByVal dwFlags As Long, _
                        ByVal dwExtraInfo As Long)

Public Type RegGraficaPrimos
  Numero As Long
  Primo As Integer
  CX As Double
  CY As Double
  Tamaño As Integer
  Color As Integer
  PCX As Double
  PCY As Double
End Type

Public rGraficaPrimos As RegGraficaPrimos

Public cn As ADODB.Connection
Public sql As String

Public EX As Double
Public EY As Double

' Calcula la distancia entre dos puntos
Public Function Distancia(ByVal pX As Double, ByVal pY As Double) As Double
  Distancia = Sqr((pX ^ 2) + (pY ^ 2))
End Function

' Transforma X
Public Function proyX(ByVal pX As Double, ByVal pY As Double) As Double
  proyX = ((4 * pX) / (4 + (Distancia(pX, pY) ^ 2)))
End Function

' Transforma Y
Public Function proyY(ByVal pX As Double, ByVal pY As Double) As Double
  proyY = ((4 * pY) / (4 + (Distancia(pX, pY) ^ 2)))
End Function

' Transforma Z
Public Function proyZ(ByVal pX As Double, ByVal pY As Double) As Double
'proyZ = 300 + ((Distancia(pX, pY) ^ 2) - 1) / (1 + (Distancia(pX, pY) ^ 2))
  proyZ = 1 - (8 / (Distancia(pX, pY) + 4))
End Function


