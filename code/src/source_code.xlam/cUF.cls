VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Listar() As Collection

    Dim col As New Collection

    col.Add "AC;Acre"
    col.Add "AL;Alagoas"
    col.Add "AP;Amapá"
    col.Add "AM;Amazonas"
    col.Add "BA;Bahia"
    col.Add "CE;Ceará"
    col.Add "DF;Distrito Federal"
    col.Add "ES;Espírito Santo"
    col.Add "GO;Goiás"
    col.Add "MA;Maranhão"
    col.Add "MT;Mato Grosso"
    col.Add "MS;Mato Grosso do Sul"
    col.Add "MG;Minas Gerais"
    col.Add "PA;Pará"
    col.Add "PB;Paraíba"
    col.Add "PR;Paraná"
    col.Add "PE;Pernambuco"
    col.Add "PI;Piauí"
    col.Add "RJ;Rio de Janeiro"
    col.Add "RN;Rio Grande do Norte"
    col.Add "RS;Rio Grande do Sul"
    col.Add "RO;Rondônia"
    col.Add "RR;Roraima"
    col.Add "SC;Santa Catarina"
    col.Add "SP;São Paulo"
    col.Add "SE;Sergipe"
    col.Add "TO;Tocantins"
    
    Set Listar = col

End Function
