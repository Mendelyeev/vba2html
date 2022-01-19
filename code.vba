'---------------------------------------------
'Módulo para exportar os dados da base em html
'---------------------------------------------
Option Explicit
Public quantidadeprocessos As Integer


Sub gerarhtml()
    If Not IsLoaded("z_aguarde2") Then
    z_aguarde2.Show
    z_aguarde2.TextBox2.Width = 138
    z_aguarde2.TextBox4.Width = 138
    z_aguarde2.TextBox6.Width = 138
    z_aguarde2.TextBox8.Width = 138
    End If
    z_aguarde2.TextBox10.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox10.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox10.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 65%"
    escreverhtml
    Delay
    z_aguarde2.TextBox12.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox12.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox12.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 70%"
    escreverconcluidos
    z_aguarde2.TextBox14.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox14.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox14.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 75%"
    escreverincluidos
    z_aguarde2.TextBox16.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox16.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox16.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 80%"
    escreversemana
    z_aguarde2.TextBox18.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox18.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox18.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 85%"
    escreverhoje
    z_aguarde2.TextBox20.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox20.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox20.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 90%"
    escreverpontos
    Delay
    z_aguarde2.TextBox22.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox22.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox22.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 90%"
    escrevertabela
    Delay
    z_aguarde2.TextBox24.Width = 0.2 * 138
    Delay
    z_aguarde2.TextBox24.Width = 0.6 * 138
    Delay
    z_aguarde2.TextBox24.Width = 1 * 138
    Delay
    z_aguarde2.Caption = "Salvando a base de dados ... 100%"
    escreveracordos
    UserForm9.Label2.Width = 288
    Delay
    'comandos para PesquisaWEB
    escreverlog
    'escreverhtmlsimples 'cria o arquivo index.html (so com o painel e a base de dados)
    Unload z_aguarde2
End Sub

Sub escreverhtml()
    Dim TextString As Variant, txt, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\index2.html": salvar = ThisWorkbook.Path & "\Web\index3.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    
    'classificação da base ativa em data da sentença
    Sheets("Processos em Andamento").Select
    ThisWorkbook.Sheets("Processos em Andamento").Range("Tabela1[[#Headers],[Fim Projetado]]").Select
    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Add2 Key:=Range("Tabela1[[#Headers],[#Data],[Fim Projetado]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("entrada").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("sentença").Visible = True
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("processo").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("conclusão").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("atualização").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("hoje").Visible = False
    '
    
    'classificação da tabela
    sht.Select
    sht.Range("DatasSentenças[[#Headers],[Processo]]").Select
    With ActiveWorkbook.Worksheets("_parametros").ListObjects("DatasSentenças").Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("DatasSentenças[[#All],[Data prevista Sentença]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
    'Data da atualização
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    
    TextString(34 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "grid" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">Painel de Controle</font></span>"
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluidos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B209").Value, "00") & "] Proc. Incluidos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    
    'Total de processos
    TextString(107 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & sht.Range("b21").Value & "</h1>"
    
    'Baseline
    TextString(113 - 1) = "<span class=" & Chr(34) & "text-danger" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("c19").Value & "</span>"
    TextString(114 - 1) = "<span class=" & Chr(34) & "text-muted" & Chr(34) & ">Incluidos | </span>"
    TextString(115 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("b30").Value & "</span>"
    TextString(116 - 1) = "<span class=" & Chr(34) & "text-muted" & Chr(34) & ">Concluidos (BL)</span>"
    
    'Em andamento
    TextString(126 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & sht.Range("b28").Value & "</h1>"
    
    'Suspensos
    TextString(130 - 1) = "<span class=" & Chr(34) & "text-danger" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("e29").Value & "</span>"
      
    'valores
    TextString(141 - 1) = "<h5 class=" & Chr(34) & "card-title mb-4" & Chr(34) & ">Valor Total atualizado em: " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h5>"
    TextString(142 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & Format(sht.Range("c27").Value / 1000000000, "R$ 0.00 Bi") & "</h1>"
    TextString(146 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & Format(sht.Range("B27").Value / 1000000000, "R$ 0.00 Bi") & "</span>"
    
    'avanço físico
    TextString(157 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & Format(sht.Range("B31").Value, "##.00%") & "</h1>"
    TextString(161 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i> " & Format(sht.Range("B32").Value, "##.00%") & "</span><span class=" & Chr(34) & "text-muted" & Chr(34) & "> Judicial | </span><span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & Format(sht.Range("B33").Value, "##.00%") & "</span><span class=" & Chr(34) & "text-muted" & Chr(34) & "> Arbitral</span>"
    
    'atualização
    TextString(215 - 1) = "<h6 class=" & Chr(34) & "card-subtitle text-muted" & Chr(34) & ">Valores atualizados em " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h6>"
    TextString(229 - 1) = "<h6 class=" & Chr(34) & "card-subtitle text-muted" & Chr(34) & ">Valores atualizados em " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h6>"
    
    'data prevista sentença
    TextString(251 - 1) = "<th class=" & Chr(34) & "d-none d-xl-table-cell" & Chr(34) & ">Data Prevista Sentença</th>"
           
    '10 maiores processos
    Dim i As Integer, rge1, rge2, rge3, rge4, rge5, rge6 As String
                      
            i = 146
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(261 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(262 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(263 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(264 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(265 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(266 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                               
            i = 147
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            TextString(269 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(270 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(271 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(272 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(273 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(274 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                   
            i = 148
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(277 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(278 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(279 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(280 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(281 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(282 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                   
            i = 149
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(285 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(286 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(287 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(288 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(289 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(290 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
            
            i = 150
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(293 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(294 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(295 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(296 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(297 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(298 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 151
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(301 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(302 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(303 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(304 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(305 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(306 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 152
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(309 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(310 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(311 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(312 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(313 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(314 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 153
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(317 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(318 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(319 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(320 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(321 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(322 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 154
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(325 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(326 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(327 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(328 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(329 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(330 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 155
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(333 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(334 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(335 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(336 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(337 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(338 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            '=============================
            '19/2/2021 - Aterações feitas
            '=============================
            
            'labels do rundown - Físico
            TextString(384 - 1) = "labels: ["
            For i = 131 To 155 '**ALTERAR MENSAL** nesse caso, ao mudar o mes, aumentar o valor de "i" em "i + 1",
            'ou seja, 'i = 123 (+1) to 146 (+1).
            'O gráfico sempre vai ficar no "meio!" Não precisa mudar no financeiro
            
            TextString(384 - 1) = TextString(384 - 1) & Chr(34) & sht.Cells(1, i).Value & Chr(34) & ","
            Next i
            TextString(384 - 1) = Left(TextString(384 - 1), Len(TextString(384 - 1)) - 1)
            TextString(384 - 1) = TextString(384 - 1) & "],"
            
            'rundown baseline
            For i = 392 To 415
            TextString(i - 1) = sht.Cells(6, i - 261).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(415 - 1) = Replace(TextString(415 - 1), ",", "")
            
            'rundown previsto
            For i = 423 To 446
            TextString(i - 1) = sht.Cells(3, i - 292).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(446 - 1) = Replace(TextString(446 - 1), ",", "")

            'rundown real
            For i = 454 To 464
            TextString(i - 1) = sht.Cells(9, i - 323).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(464 - 1) = Replace(TextString(464 - 1), ",", "")
          
            'top10 em valor
            TextString(514 - 1) = "labels: [" & _
            Chr(34) & sht.Range("K91").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K92").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K93").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K94").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K95").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K96").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K97").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K98").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K99").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K100").Value & Chr(34) & "],"

            TextString(521 - 1) = "data: [" & _
            Replace(Format(sht.Range("I91").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I92").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I93").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I94").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I95").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I96").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I97").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I98").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I99").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I9100").Value / 1000000000, "0.00"), ",", ".") & "],"
            
            'top 10 empreendimentos
            TextString(559 - 1) = "labels: [" & _
            Chr(34) & sht.Range("I108").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I109").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I110").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I111").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I112").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I113").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I114").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I115").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I116").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I117").Value & Chr(34) & "],"
            
            TextString(566 - 1) = "data: [" & _
            Replace(Format(sht.Range("J108").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J109").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J110").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J111").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J112").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J113").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J114").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J115").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J116").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J117").Value, "0.00"), ",", ".") & "],"
            
            'pizza processos
            TextString(649 - 1) = "labels: [" & _
            Chr(34) & sht.Range("A22").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("A23").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("A24").Value & Chr(34) & "],"
            
            TextString(651 - 1) = "data: [" & _
            sht.Range("c22").Value & ", " & _
            sht.Range("c23").Value & ", " & _
            sht.Range("c24").Value & "], "

            'labels do rundown FINANCEIRO
            TextString(687 - 1) = TextString(384 - 1) 'só precisa alterar no físico!!!
            
            'Financeiro
            'rundown baseline
            For i = 695 To 718
            TextString(i - 1) = Replace(Format(sht.Cells(18, i - 564).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(718 - 1) = Replace(TextString(718 - 1), ",", "")
            
            'rundown previsto
            For i = 726 To 749
            TextString(i - 1) = Replace(Format(sht.Cells(14, i - 595).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(749 - 1) = Replace(TextString(749 - 1), ",", "")

            'rundown real
            For i = 757 To 767
            TextString(i - 1) = Replace(Format(sht.Cells(16, i - 626).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(767 - 1) = Replace(TextString(767 - 1), ",", "")

            'retirada dos caracteres especiais
            For i = 251 To 570
            TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
            TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
            TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
            TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
            TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
            TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
            TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
            TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
            TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
            TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
            TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
            TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
            TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
            Next i
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
End Sub
Sub escreverconcluidos()

    Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, rge10, rge11, rge12, rge13, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html": salvar = ThisWorkbook.Path & "\Web\concluidos.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
    ThisWorkbook.Sheets("Processos em Andamento").Select
    sortcon
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    texto = "<tt>" & texto & vbNewLine + "----------------------------------------</br>"
    texto = texto + vbNewLine + "Total da Carteira PPGSP/AC_________[" & ThisWorkbook.Sheets("_parametros").Range("B21").Value & "]</br>"
    texto = texto + vbNewLine + "Total de Processos (em andamento)__[" & ThisWorkbook.Sheets("_parametros").Range("B28").Value & "]</br>"
    texto = texto + vbNewLine + "Processos Concluídos_______________[" & Format(ThisWorkbook.Sheets("_parametros").Range("B30").Value, "000") & "]</br>"
    texto = texto + vbNewLine + "----------------------------------------</br>"
    texto = texto + vbNewLine + "Base de dados que gerou esse arquivo: " & ThisWorkbook.Name
    texto = texto & vbNewLine & "</DIV>"
    texto = texto & vbNewLine & "</DIV>"
         
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Processos em Andamento")
        For j = 3 To sht2.Range("b3", sht2.Range("B3").End(xlDown)).Rows.Count + 2
            RGE = "B" & j
            rge1 = "U" & j
            rge2 = "AS" & j
            rge3 = "G" & j
            rge4 = "AR" & j
            rge5 = "AG" & j
            rge6 = "S" & j
            rge7 = "I" & j
            rge8 = "M" & j
            rge9 = "V" & j
            rge10 = "Z" & j
            rge11 = "R" & j
            rge12 = "AE" & j
            rge13 = "AD" & j
            If (sht2.Range(rge1).Value = "[CONCLUÍDO]") And (sht2.Range(rge2).Value <> "AJ") Then
            i = i + 1
            i = Format(i, "00")
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "justificado" & Chr(34) & "><span style=" & Chr(34) & "color:#00b0f0" & Chr(34) & ">"
            texto = texto & vbNewLine & "[" & i & "] " & sht2.Range(RGE).Value & "</br></span>"
            texto = texto & vbNewLine & "-> Contraparte:       <u>" & sht2.Range(rge3).Value & "</u></br>"
            texto = texto & vbNewLine & "-> Polo:              <u>" & sht2.Range(rge8).Value & "</u></br>"
            texto = texto & vbNewLine & "-> Ponto Focal:       " & sht2.Range(rge7).Value & "</br>"
            texto = texto & vbNewLine & "-> Data da conclusão: [" & sht2.Range(rge4).Value & "]</br>"
            texto = texto & vbNewLine & "-> Valores (atualizado | P(0)): [" & Format(sht2.Range(rge10).Value, "R$ #,###,###,##0.00") & " | " & Format((sht2.Range(rge12).Value - sht2.Range(rge13).Value), "R$ #,###,###,##0.00") & "]</br>"
            texto = texto & vbNewLine & "-> Valor da Sentença/Acordo: [<u><b>" & Format(sht2.Range(rge5).Value, "R$ #,###,###,##0.00") & "</b></u>]</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "<u>Última observação registrada na base [Atualizado em: " & Format(sht2.Range(rge9).Value, "dd/mm/yyyy") & "]:</u></br>"
            texto = texto & vbNewLine & sht2.Range(rge6).Value & "</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "<u>Pedidos (Valores em R$ Milhões e a P[0]):</u></br>"
            texto = texto & vbNewLine & sht2.Range(rge11).Value & "</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
            End If
            
        Next j
    
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(37 - 1) = "<li class=" & Chr(34) & " sidebar-item active" & Chr(34) & ">"
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "><font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</font></span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B209").Value, "00") & "] Proc. Incluídos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Processos Finalizados (até emissão da Sentença - 1.a Instância):</h1>"
    TextString(102 - 1) = texto & "</tt>"
    
        For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
End Sub
Sub escreverincluidos()

    Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, rge10, rge11, rge12, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html": salvar = ThisWorkbook.Path & "\Web\novos.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
    ThisWorkbook.Sheets("Processos em Andamento").Select
    Range("Tabela1[[#Headers],[Data de Entrada SRGE/PPSGP/AC]]").Select

    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Add2 Key:=Range( _
        "Tabela1[[#Headers],[#Data],[Data de Entrada SRGE/PPSGP/AC]]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    texto = vbNewLine + "<tt>------------------------------------------</br>"
    texto = texto + vbNewLine + "Total da Carteira PPGSP/AC___________[" & sht.Range("B21").Value & "]</br>"
    texto = texto + vbNewLine + "Total de Processos (em andamento)____[" & sht.Range("B28").Value & "]</br>"
    texto = texto + vbNewLine + "Processos incluídos desde " & sht.Range("B19").Value & "_[" & Format(sht.Range("c19").Value, "000") & "]</br>"
    texto = texto + vbNewLine + "------------------------------------------</br>"
    texto = texto + vbNewLine + "Base de dados que gerou esse arquivo: " & ThisWorkbook.Name
    texto = texto & vbNewLine & "</DIV>"
    texto = texto & vbNewLine & "</DIV>"
         
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Processos em Andamento")
        For j = 3 To sht2.Range("b3", sht2.Range("B3").End(xlDown)).Rows.Count + 2
            RGE = "B" & j
            rge1 = "U" & j
            rge2 = "Z" & j
            rge3 = "G" & j
            rge4 = "AR" & j
            rge5 = "AS" & j
            rge6 = "S" & j
            rge7 = "I" & j
            rge8 = "M" & j
            rge9 = "AQ" & j
            rge10 = "R" & j
            rge11 = "AE" & j
            rge12 = "AD" & j
            If (sht2.Range(rge9).Value >= 43922) And (sht2.Range(rge5).Value <> "AJ") Then
            i = i + 1
            i = Format(i, "00")
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "justificado" & Chr(34) & "><span style=" & Chr(34) & "color:red" & Chr(34) & ">"
            texto = texto & vbNewLine & "[" & i & "] " & sht2.Range(RGE).Value & "</br></span>"
            texto = texto & vbNewLine & "Data de entrada:      " & Format(sht2.Range(rge9).Value, "dd/mm/yyyy") & "</br>"
            texto = texto & vbNewLine & "-> Contraparte:       <u>" & sht2.Range(rge3).Value & "</u></br>"
            texto = texto & vbNewLine & "-> Polo:              <u>" & sht2.Range(rge8).Value & "</u></br>"
            texto = texto & vbNewLine & "-> Ponto Focal:       " & sht2.Range(rge7).Value & "</br>"
            texto = texto & vbNewLine & "-> Status:       " & sht2.Range(rge1).Value & "</br>"
            texto = texto & vbNewLine & "-> Valores (atualizado | P(0)):       <u><b>" & Format(sht2.Range(rge2).Value, "R$ #,###,###,000.00") & " | " & Format((sht2.Range(rge11).Value - sht2.Range(rge12).Value), "R$ #,###,###,000.00") & "</b></u></br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "<u>Última observação registrada na base:</u></br>"
            texto = texto & vbNewLine & sht2.Range(rge6).Value & "</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "<u>Pedidos (Valores em R$ Milhões e a P[0]):</u></br>"
            texto = texto & vbNewLine & sht2.Range(rge10).Value & "</br>"
            texto = texto & vbNewLine & "------------------------------------------</br>"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
            End If
        Next j
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(42 - 1) = "<li class=" & Chr(34) & " sidebar-item active" & Chr(34) & ">"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</font></span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Processos Incluídos após 01/04/2020 (data de congelamento da baseline):</h1>"
    TextString(102 - 1) = texto & "</tt>"
    
        For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
End Sub
Sub escreverhoje()
Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html": salvar = ThisWorkbook.Path & "\Web\hoje.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Processos em Andamento")
        texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
        texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & "><tt>"
        For j = 3 To sht2.Range("b3", sht2.Range("B3").End(xlDown)).Rows.Count + 2
            RGE = "B" & j
            rge1 = "Y" & j
            rge2 = "v" & j
            rge3 = "G" & j
                If sht2.Range(rge2).Value = Date Then
                    i = i + 1
                    i = Format(i, "00")
                    texto = texto & vbNewLine & "[" & i & "] " & sht2.Range(RGE).Value & " (" & sht2.Range(rge3).Value & ")</br>"
                End If
        Next j
        texto = texto & vbNewLine & "</DIV>"
        texto = texto & vbNewLine & "</DIV></tt>"
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</span>"
    TextString(47 - 1) = "<li class=" & Chr(34) & " sidebar-item active" & Chr(34) & ">"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</font></span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(97 - 1) = vbNullString
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Processos verificados e/ou atualizados hoje:</h1>"
    TextString(101 - 1) = vbNullString
    TextString(102 - 1) = texto
    
    For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
   
End Sub
Sub escreversemana()
Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html": salvar = ThisWorkbook.Path & "\Web\semana.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Processos em Andamento")
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & "><TT>"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & ">"
            texto = texto & vbNewLine & "Semana Atual: [" & ThisWorkbook.Sheets("_parametros").Range("B153").Value & ".ª]</br> [De:" & ThisWorkbook.Sheets("_parametros").Range("B151").Value & " a: " & ThisWorkbook.Sheets("_parametros").Range("B152").Value & "]"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & ">"
        For j = 3 To sht2.Range("b3", sht2.Range("B3").End(xlDown)).Rows.Count + 2
            RGE = "B" & j
            rge1 = "Y" & j
            rge2 = "AN" & j
            rge3 = "G" & j
            If sht2.Range(rge1).Value <> sht2.Range(rge2).Value Then
            i = i + 1
            i = Format(i, "00")
            texto = texto & vbNewLine & "--------------------------------------------</BR>"
            texto = texto & vbNewLine & "[" & i & "] " & sht2.Range(RGE).Value & " (" & sht2.Range(rge3).Value & ")</br>"
            texto = texto & vbNewLine & "De: [<U>" & sht2.Range(rge2).Value & "</U>] para: [<u>" & sht2.Range(rge1).Value & "</u>]</br>"
            texto = texto & vbNewLine & "--------------------------------------------</BR>"
            End If
        Next j
           texto = texto & vbNewLine & "</DIV>"
           texto = texto & vbNewLine & "</DIV></TT>"

    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(52 - 1) = "<li class=" & Chr(34) & " sidebar-item active" & Chr(34) & ">"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</font></span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(97 - 1) = vbNullString
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Processos com avanço na semana atual:</h1>"
    TextString(101 - 1) = vbNullString
    TextString(102 - 1) = texto
    
        For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
       
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
   
End Sub
Sub escreverpontos()
Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html": salvar = ThisWorkbook.Path & "\Web\pontos.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Controle de Recebimento")
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
            texto = texto & vbNewLine & "<div class=" & Chr(34) & "card-body" & Chr(34) & "><tt>"
            texto = texto & vbNewLine & "Prazo: 6.º dia do mês"
            texto = texto & vbNewLine & "</DIV>"
            texto = texto & vbNewLine & "</DIV>"
        
        texto = texto & vbNewLine & "<div class=" & Chr(34) & "card" & Chr(34) & ">"
        For j = 3 To ThisWorkbook.Sheets("Controle de Recebimento").Range("f3", sht2.Range("f3").End(xlDown)).Rows.Count + 2
            RGE = "F" & j
            rge1 = "I" & j
            If sht2.Range(rge1).Value = True Then
            texto = texto & vbNewLine & "<font color=" & Chr(34) & "#04cc15" & Chr(34) & ">[OK] Ponto Focal: [" & sht2.Range(RGE).Value & "]</font></br>"
            Else
            texto = texto & vbNewLine & "<font color=" & Chr(34) & "#ff0000" & Chr(34) & ">Ponto Focal: [" & sht2.Range(RGE).Value & "] - PENDENTE </font></br>"
            End If
        Next j
        texto = texto & vbNewLine & "</tt></DIV>"

    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(55 - 1) = "</a><li class=" & Chr(34) & "sidebar-item active" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pontos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "user-check" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">Pts. Focais - Update</font></span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "acordos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "thumbs-up" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Acordos</font></span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pesquisa.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "search" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pesquisa Processual</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pendencias.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "alert-triangle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pendências</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "log.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "watch" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Log</span></a>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(97 - 1) = vbNullString
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Status das Atualizações dos Pontos Focais</h1>"
    TextString(101 - 1) = vbNullString
    TextString(102 - 1) = texto
    
        For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
   
End Sub
Sub escrevertabela()
Dim TextString As Variant, texto, abrir, salvar As String, salvar2 As String
    abrir = ThisWorkbook.Path & "\Web\pesquisabase.html"
    salvar = ThisWorkbook.Path & "\Web\pesquisa.aspx"
    
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
        
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
    'script para a transferência da Tabela1 para Array no HTML:
    Dim sht2 As Object, i, j As Integer
    Set sht2 = ThisWorkbook.Sheets("Processos em Andamento")
    
    texto = "<script>var myArray = ["
    For j = 3 To sht2.Range("B3", sht2.Range("B3").End(xlDown)).Rows.Count + 2
        texto = texto & "{"
            i = 2
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), " "), "'", " "), "\", "/"), 100) & "', "
            i = 3
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 5
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 7
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 9
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 11
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 13
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 16
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 17
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500), "dd/mm/yyyy") & "', "
            i = 18
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), ""), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500) & "', "
            i = 19
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), ""), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500) & "', "
            i = 20
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), ""), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500) & "', "
            i = 21
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), ""), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500) & "', "
            i = 23
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500), "0.00%") & "', "
            i = 25
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500) & "', "
            i = 26
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 500), "R$ #,###,###,##0.00") & "', "
            i = 33
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", ""), "\", "/"), 500), "R$ #,###,###,##0.00") & "', "
            For i = 35 To 38
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", ""), "\", "/"), 500), "R$ #,###,###,##0.00") & "', "
            Next i
            i = 42
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Format(Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), ""), Chr(13), ""), Chr(13) & Chr(10), ""), "'", ""), "\", "/"), 500), "R$ #,###,###,##0.00") & "', "
            i = 46
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 47
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 49
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
            i = 50
            texto = texto & "'" & Replace(Replace(Replace(Replace(Replace(sht2.Cells(2, i).Value, " ", "_"), "(", ""), ")", ""), "/", ""), "-", "") & "':'" & Left(Replace(Replace(Replace(Replace(Replace(sht2.Cells(j, i).Value, Chr(10), " "), Chr(13), " "), Chr(13) & Chr(10), ""), "'", " "), "\", "/"), 100) & "', "
        texto = Left(texto, Len(texto) - 1)
        texto = texto & vbNewLine
        texto = texto & "},"
    Next j
    texto = texto & "]</script>"
            texto = Replace(texto, "ç", "c")
            texto = Replace(texto, "ç", "c")
            texto = Replace(texto, "ã", "a")
            texto = Replace(texto, "õ", "o")
            texto = Replace(texto, "á", "a")
            texto = Replace(texto, "é", "e")
            texto = Replace(texto, "í", "i")
            texto = Replace(texto, "ó", "o")
            texto = Replace(texto, "ú", "u")
            texto = Replace(texto, "â", "a")
            texto = Replace(texto, "ê", "e")
            texto = Replace(texto, "ô", "o")
            texto = Replace(texto, "Ç", "C")
            texto = Replace(texto, "Ã", "A")
            texto = Replace(texto, "Õ", "O")
            texto = Replace(texto, "Á", "A")
            texto = Replace(texto, "É", "E")
            texto = Replace(texto, "Í", "I")
            texto = Replace(texto, "Ó", "O")
            texto = Replace(texto, "Ú", "U")
            texto = Replace(texto, "Â", "A")
            texto = Replace(texto, "Ê", "E")
            texto = Replace(texto, "Ô", "O")
            texto = Replace(texto, "º", "_")
            texto = Replace(texto, "Contratada", "contratada")
      
    TextString(127 - 1) = texto
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(10)))

    
End Sub
Sub escreveracordos()
Dim TextString As Variant, RGE, rge1, rge2, rge3, rge4, rge5, rge6, rge7, rge8, rge9, texto, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\acordosbase.html": salvar = ThisWorkbook.Path & "\Web\acordos.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    'otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
        Dim sht2 As Worksheet, i, j As Integer
        Set sht2 = ThisWorkbook.Sheets("Acordos em Andamento")
                      
        For j = 47 To 61 'aumetar se incluir novos acordos (1.a linha de dados da tabela até a última linha de dados)
            rge1 = "A" & j
            rge2 = "B" & j
            rge3 = "C" & j
            rge4 = "D" & j
            rge5 = "E" & j
            rge6 = "F" & j
            texto = texto & "<tr>" & vbNewLine
            If sht2.Range(rge6).Value = "9. Concluído" Then
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-trans" & Chr(34) & ">" & sht2.Range(rge1).Value & "</span></td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge2).Value, "R$ #,###,###,###.00") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge3).Value & "</td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge4).Value, "dd/mm/yyyy") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge5).Value & "</td>"
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht2.Range(rge6).Value & "</span></td>"
            ElseIf sht2.Range(rge6).Value = "CANCELADO" Then
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-trans" & Chr(34) & ">" & sht2.Range(rge1).Value & "</span></td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge2).Value, "R$ #,###,###,###.00") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge3).Value & "</td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge4).Value, "dd/mm/yyyy") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge5).Value & "</td>"
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-amarelo" & Chr(34) & ">" & sht2.Range(rge6).Value & "</span></td>"
            Else
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-trans" & Chr(34) & ">" & sht2.Range(rge1).Value & "</span></td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge2).Value, "R$ #,###,###,###.00") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge3).Value & "</td>"
                texto = texto & vbNewLine & "<td>" & Format(sht2.Range(rge4).Value, "dd/mm/yyyy") & "</td>"
                texto = texto & vbNewLine & "<td>" & sht2.Range(rge5).Value & "</td>"
                texto = texto & vbNewLine & "<td><span class=" & Chr(34) & "badge badge-azul" & Chr(34) & ">" & sht2.Range(rge6).Value & "</span></td>"
            End If
            texto = texto & vbNewLine & "</tr>" & vbNewLine
        Next j
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(55 - 1) = "</a><li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pontos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "user-check" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pts. Focais - Update</a>" _
    & "<li class=" & Chr(34) & "sidebar-item active" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "acordos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "thumbs-up" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "><font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">Acordos</font></span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pesquisa.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "search" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pesquisa Processual</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pendencias.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "alert-triangle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pendências</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "log.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "watch" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Log</span></a>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Situção atual dos Acordos:</h1>"
    TextString(110 - 1) = texto
    
    For i = 1 To 110
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(10)))
    'otimizaOFF
   
End Sub

Sub escreverlog()
Dim TextString, TextString2 As Variant, texto, RGE, log, abrir, salvar As String, sht As Object, i As Integer
    abrir = ThisWorkbook.Path & "\Web\pages-blank.html"
    salvar = ThisWorkbook.Path & "\Web\log.aspx"
    log = ThisWorkbook.Path & "\" & "(!) Base Consolidada - Status e Log " & Format(Date, "yyyy-mm-dd") & ".txt"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    texto = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(log).ReadAll, Chr(10))
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    

    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluídos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Right(0 & ThisWorkbook.Sheets("_parametros").Range("B209").Value, 2) & "] Proc. Incluídos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(55 - 1) = "</a><li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pontos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "user-check" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> Pts. Focais - Update</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "acordos.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "thumbs-up" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Acordos</font></span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pesquisa.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "search" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pesquisa Processual</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pendencias.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "alert-triangle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pendências</span></a>" _
    & "<li class=" & Chr(34) & "sidebar-item active" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "log.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "watch" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "><font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">Log</font></span></a>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    TextString(93 - 1) = "<h1 class=" & Chr(34) & "h3 mb-3" & Chr(34) & ">Log da Base de Dados</h1>"
    TextString(97 - 1) = "<div class=" & Chr(34) & "card" & Chr(34) & "><div class=" & Chr(34) & "card-body" & Chr(34) & "><tt><p style=" & Chr(34) & "font-size:12px" & Chr(34) & ">"
    For i = 0 To UBound(texto)
    texto(i) = texto(i) & "<br>"
    Next
    TextString(98 - 1) = Join(texto) & "</div></div></tt></p>"
    TextString(101 - 1) = vbNullString
    
        For i = 1 To 102
        TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
        TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
        TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
        TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
        TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
        TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
        TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
        TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
        TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
        TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
        TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
        TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
        TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
    Next
    
    'write back in file
    CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
End Sub


Sub escreverhtmlsimples()
    Dim TextString As Variant, txt, abrir, salvar As String, sht As Object
    abrir = ThisWorkbook.Path & "\Web\index2.html": salvar = "pathAC\index.aspx"
    Set sht = ThisWorkbook.Sheets("_parametros")
    
    'otimização do código
    otimizaON
    
    'classificação da base ativa em data da sentença
    Sheets("Processos em Andamento").Select
    ThisWorkbook.Sheets("Processos em Andamento").Range("Tabela1[[#Headers],[Fim Projetado]]").Select
    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1").Sort _
        .SortFields.Add2 Key:=Range("Tabela1[[#Headers],[#Data],[Fim Projetado]]"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Processos em Andamento").ListObjects("Tabela1") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("entrada").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("sentença").Visible = True
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("processo").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("conclusão").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("atualização").Visible = False
    ThisWorkbook.Sheets("Processos em Andamento").Shapes("hoje").Visible = False
    '
    
    'classificação da tabela
    sht.Select
    sht.Range("DatasSentenças[[#Headers],[Processo]]").Select
    With ActiveWorkbook.Worksheets("_parametros").ListObjects("DatasSentenças").Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range("DatasSentenças[[#All],[Data prevista Sentença]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    If FileExists(salvar) = True Then
      Kill salvar
    Else
    End If
      'read text from file
    TextString = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile(abrir).ReadAll, Chr(10))
    
    'Data da atualização
    TextString(29 - 1) = "Data: " & Format(sht.Range("B1").Value, "dd/mm/yyyy") & " - hora: " & Format(sht.Range("D1").Value, "hh:mm:ss")
    
    TextString(34 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "grid" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & "> <font color=" & Chr(34) & "#f4fc49" & Chr(34) & ">Painel de Controle</font></span>"
    TextString(39 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "check-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & ThisWorkbook.Sheets("_parametros").Range("B30").Value & "] Proc. Concluidos</span>"
    TextString(44 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "plus-circle" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B209").Value, "00") & "] Proc. Incluidos</span>"
    TextString(49 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "eye" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("Processos em Andamento").Range("I1").Value, "00") & "] Proc. Verif. Hoje</span>"
    TextString(54 - 1) = "<i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "clock" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">[" & Format(ThisWorkbook.Sheets("_parametros").Range("B150").Value, "00") & "] Proc. C/mov. Sem.</span>"
    TextString(70 - 1) = "<a href=" & Chr(34) & ThisWorkbook.Name & Chr(34) & " download><button class=" & Chr(34) & "btn" & Chr(34) & " type=" & Chr(34) & "button" & Chr(34) & ">"
    
    'Total de processos
    TextString(107 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & sht.Range("b21").Value & "</h1>"
    
    'Baseline
    TextString(113 - 1) = "<span class=" & Chr(34) & "text-danger" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("c19").Value & "</span>"
    TextString(114 - 1) = "<span class=" & Chr(34) & "text-muted" & Chr(34) & ">Incluidos | </span>"
    TextString(115 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("b30").Value & "</span>"
    TextString(116 - 1) = "<span class=" & Chr(34) & "text-muted" & Chr(34) & ">Concluidos (BL)</span>"
    
    'Em andamento
    TextString(126 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & sht.Range("b28").Value & "</h1>"
    
    'Suspensos
    TextString(130 - 1) = "<span class=" & Chr(34) & "text-danger" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & sht.Range("e29").Value & "</span>"
      
    'valores
    TextString(141 - 1) = "<h5 class=" & Chr(34) & "card-title mb-4" & Chr(34) & ">Valor Total atualizado em: " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h5>"
    TextString(142 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & Format(sht.Range("c27").Value / 1000000000, "R$ 0.00 Bi") & "</h1>"
    TextString(146 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & Format(sht.Range("B27").Value / 1000000000, "R$ 0.00 Bi") & "</span>"
    
    'avanço físico
    TextString(157 - 1) = "<h1 class=" & Chr(34) & "display-5 mt-1 mb-3" & Chr(34) & ">" & Format(sht.Range("B31").Value, "##.00%") & "</h1>"
    TextString(161 - 1) = "<span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i> " & Format(sht.Range("B32").Value, "##.00%") & "</span><span class=" & Chr(34) & "text-muted" & Chr(34) & "> Judicial | </span><span class=" & Chr(34) & "text-success" & Chr(34) & "> <i class=" & Chr(34) & "mdi mdi-arrow-bottom-right" & Chr(34) & "></i>" & Format(sht.Range("B33").Value, "##.00%") & "</span><span class=" & Chr(34) & "text-muted" & Chr(34) & "> Arbitral</span>"
    
    'atualização
    TextString(215 - 1) = "<h6 class=" & Chr(34) & "card-subtitle text-muted" & Chr(34) & ">Valores atualizados em " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h6>"
    TextString(229 - 1) = "<h6 class=" & Chr(34) & "card-subtitle text-muted" & Chr(34) & ">Valores atualizados em " & Format(ThisWorkbook.Sheets("Estratificacoes_AC_Consolidado").Range("c1").Value, "dd/mm/yyyy") & "</h6>"
    
    'data prevista sentença
    TextString(251 - 1) = "<th class=" & Chr(34) & "d-none d-xl-table-cell" & Chr(34) & ">Data Prevista Sentença</th>"
           
    '10 maiores processos
    Dim i As Integer, rge1, rge2, rge3, rge4, rge5, rge6 As String
                      
            i = 146
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(261 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(262 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(263 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(264 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(265 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(266 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                               
            i = 147
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            TextString(269 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(270 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(271 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(272 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(273 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(274 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                   
            i = 148
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(277 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(278 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(279 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(280 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(281 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(282 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
                   
            i = 149
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(285 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(286 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(287 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(288 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(289 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(290 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"
            
            i = 150
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(293 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(294 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(295 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(296 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(297 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(298 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 151
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(301 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(302 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(303 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(304 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(305 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(306 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 152
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(309 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(310 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(311 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(312 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(313 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(314 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 153
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(317 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(318 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(319 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(320 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(321 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(322 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 154
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(325 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(326 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(327 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(328 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(329 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(330 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            i = 155
            rge1 = "I" & i
            rge2 = "K" & i
            rge3 = "L" & i
            rge4 = "M" & i
            rge5 = "N" & i
            rge6 = "O" & i
            
            TextString(333 - 1) = "<td>" & sht.Range(rge1).Value & "</td>"
            TextString(334 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge6).Value & "</td>"
            TextString(335 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge2).Value, "dd/mm/yyyy") & "</td>"
            TextString(336 - 1) = "<td><span class=" & Chr(34) & "badge badge-success" & Chr(34) & ">" & sht.Range(rge3).Value & "</span></td>"
            TextString(337 - 1) = "<td class=" & Chr(34) & "d-md-table-cell" & Chr(34) & ">" & sht.Range(rge4).Value & "</td>"
            TextString(338 - 1) = "<td class=" & Chr(34) & "d-xl-table-cell" & Chr(34) & ">" & Format(sht.Range(rge5).Value * 1000, "R$ 0.000 Mi") & "</td>"

            '=============================
            '19/2/2021 - Aterações feitas
            '=============================
            
            'labels do rundown - Físico
            TextString(384 - 1) = "labels: ["
            For i = 129 To 152 '**ALTERAR MENSAL** nesse caso, ao mudar o mes, aumentar o valor de "i" em "i + 1",
            'ou seja, 'i = 123 (+1) to 146 (+1).
            'O gráfico sempre vai ficar no "meio!" Não precisa mudar no financeiro
            
            TextString(384 - 1) = TextString(384 - 1) & Chr(34) & sht.Cells(1, i).Value & Chr(34) & ","
            Next i
            TextString(384 - 1) = Left(TextString(384 - 1), Len(TextString(384 - 1)) - 1)
            TextString(384 - 1) = TextString(384 - 1) & "],"
            
            'rundown baseline
            For i = 392 To 415
            TextString(i - 1) = sht.Cells(6, i - 263).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(415 - 1) = Replace(TextString(415 - 1), ",", "")
            
            'rundown previsto
            For i = 423 To 446
            TextString(i - 1) = sht.Cells(3, i - 294).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(446 - 1) = Replace(TextString(446 - 1), ",", "")

            'rundown real
            For i = 454 To 464
            TextString(i - 1) = sht.Cells(9, i - 325).Value & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(464 - 1) = Replace(TextString(464 - 1), ",", "")
          
            'top10 em valor
            TextString(514 - 1) = "labels: [" & _
            Chr(34) & sht.Range("K91").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K92").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K93").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K94").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K95").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K96").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K97").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K98").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K99").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("K100").Value & Chr(34) & "],"

            TextString(521 - 1) = "data: [" & _
            Replace(Format(sht.Range("I91").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I92").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I93").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I94").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I95").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I96").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I97").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I98").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I99").Value / 1000000000, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("I9100").Value / 1000000000, "0.00"), ",", ".") & "],"
            
            'top 10 empreendimentos
            TextString(559 - 1) = "labels: [" & _
            Chr(34) & sht.Range("I108").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I109").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I110").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I111").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I112").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I113").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I114").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I115").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I116").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("I117").Value & Chr(34) & "],"
            
            TextString(566 - 1) = "data: [" & _
            Replace(Format(sht.Range("J108").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J109").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J110").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J111").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J112").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J113").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J114").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J115").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J116").Value, "0.00"), ",", ".") & ", " & _
            Replace(Format(sht.Range("J117").Value, "0.00"), ",", ".") & "],"
            
            'pizza processos
            TextString(649 - 1) = "labels: [" & _
            Chr(34) & sht.Range("A22").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("A23").Value & Chr(34) & ", " & _
            Chr(34) & sht.Range("A24").Value & Chr(34) & "],"
            
            TextString(651 - 1) = "data: [" & _
            sht.Range("c22").Value & ", " & _
            sht.Range("c23").Value & ", " & _
            sht.Range("c24").Value & "], "

            'labels do rundown FINANCEIRO
            TextString(687 - 1) = TextString(384 - 1) 'só precisa alterar no físico!!!
            
            'Financeiro
            'rundown baseline
            For i = 695 To 718
            TextString(i - 1) = Replace(Format(sht.Cells(18, i - 566).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(718 - 1) = Replace(TextString(718 - 1), ",", "")
            
            'rundown previsto
            For i = 726 To 749
            TextString(i - 1) = Replace(Format(sht.Cells(14, i - 597).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(749 - 1) = Replace(TextString(749 - 1), ",", "")

            'rundown real
            For i = 757 To 767
            TextString(i - 1) = Replace(Format(sht.Cells(16, i - 628).Value, "0.00"), ",", ".") & "," '**ALTERAR MENSAL** ao mudar o mes, DIMINUIR o valor de "i" em "i - 1"
            Next i
            TextString(767 - 1) = Replace(TextString(767 - 1), ",", "")

            'retirada dos caracteres especiais
            For i = 251 To 570
            TextString(i - 1) = Replace(TextString(i - 1), "ç", "c")
            TextString(i - 1) = Replace(TextString(i - 1), "ã", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "õ", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "á", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "é", "e")
            TextString(i - 1) = Replace(TextString(i - 1), "í", "i")
            TextString(i - 1) = Replace(TextString(i - 1), "ó", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "ú", "u")
            TextString(i - 1) = Replace(TextString(i - 1), "â", "a")
            TextString(i - 1) = Replace(TextString(i - 1), "ê", "e")
            TextString(i - 1) = Replace(TextString(i - 1), "ô", "o")
            TextString(i - 1) = Replace(TextString(i - 1), "Ç", "C")
            TextString(i - 1) = Replace(TextString(i - 1), "Ã", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "Õ", "O")
            TextString(i - 1) = Replace(TextString(i - 1), "Á", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "É", "E")
            TextString(i - 1) = Replace(TextString(i - 1), "Í", "I")
            TextString(i - 1) = Replace(TextString(i - 1), "Ó", "O")
            TextString(i - 1) = Replace(TextString(i - 1), "Ú", "U")
            TextString(i - 1) = Replace(TextString(i - 1), "Â", "A")
            TextString(i - 1) = Replace(TextString(i - 1), "Ê", "E")
            TextString(i - 1) = Replace(TextString(i - 1), "Ô", "O")
            Next i
    'remoção dos links que não irão atualizar:'
    For i = 37 To 55
    TextString(i - 1) = ""
    Next
    For i = 67 To 74
    TextString(i - 1) = ""
    Next
    
    TextString(37 - 1) = "<li class=" & Chr(34) & "sidebar-item" & Chr(34) & "><a class=" & Chr(34) & "sidebar-link" & Chr(34) & " href=" & Chr(34) & "pesquisa.aspx" & Chr(34) & "><i class=" & Chr(34) & "align-middle" & Chr(34) & " data-feather=" & Chr(34) & "search" & Chr(34) & "></i> <span class=" & Chr(34) & "align-middle" & Chr(34) & ">Pesquisa Processual</span></a></li>"
    
    'write back in file
     CreateObject("Scripting.FileSystemObject").CreateTextFile(salvar).Write (Join(TextString, Chr(13) & Chr(10)))
    'otimizaOFF
End Sub


