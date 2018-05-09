Imports System.Drawing
Imports System.Text.RegularExpressions
Imports System.Windows.Input
Imports Microsoft.VisualStudio.TestTools.UITest.Extension
Imports Microsoft.VisualStudio.TestTools.UITesting
Imports Microsoft.VisualStudio.TestTools.UITesting.Keyboard
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Excel = Microsoft.Office.Interop.Excel

<CodedUITest()>
Public Class Gerar_Massa_de_Dados_P5
    Public Conexao As OleDb.OleDbConnection
    Dim Query_str As String
    Dim Parte_Execucao As Int32 = 1
    Public Data_Alteracao_Status As String = Year(DateTime.Now).ToString() + "-" + Month(DateTime.Now).ToString() + "-" + (Day(DateTime.Now) - 1).ToString()
    Public Data_Retorno_Coletor As String = Year(DateTime.Now).ToString() + "-" + Month(DateTime.Now).ToString() + "-" + Day(DateTime.Now).ToString()
    Public Mensagem As String
    Public Retorno, Retorno_2 As String()
    Public Quantidade_Registro_Coletor As Int32

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_2()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6"
        Dim vRows_Final As String = "596"
        Dim Unidade_Negocio_Origem As String = "VD177"
        Dim Unidade_Negocio_Destino As String = "VD910"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\VD910 . ST910 - CT02 - PR05.xlsx"
        Dim Nome_Aba_Arquivo As String = "VD910 > ST910 - CT02 - PR05"
        Dim Item_Not_Execute As String

        Rota = "05"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_DPSP_HDR_DEVO_COLETOR_SP", Data_Retorno_Coletor, "", "", "DPSP_WMS_ORDER_NO").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                Mensagem = Script_BD(1, "PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If ((Mensagem <> "") And (Quantidade_Registro_Coletor > 0)) Then
                Mensagem = ""
                Mensagem = Script_BD(1, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If ((Mensagem <> "") And (Quantidade_Registro_Coletor > 0)) Then
                Mensagem = ""
                Mensagem = Script_BD(2, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_4()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "8, 1001"
        Dim vRows_Final As String = "1000, 1226"
        Dim Unidade_Negocio_Origem As String = "VD647"
        Dim Unidade_Negocio_Destino As String = "VD910"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\VD910.xlsx"
        Dim Nome_Aba_Arquivo As String = "P5 - Saída CD"
        Dim Item_Not_Execute As String
        Rota = "60"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Retorno = Script_BD(0, "PS_DPSP_HDR_DEVO_COLETOR_SP", Data_Retorno_Coletor, "", "", "DPSP_WMS_ORDER_NO").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                Mensagem = Script_BD(1, "PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(1, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(2, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_ES()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "7,1001"
        Dim vRows_Final As String = "1000,1225"
        Dim Unidade_Negocio_Origem As String = "L1246"
        Dim Unidade_Negocio_Destino As String = "VD906"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\P5 - Saída CD.xlsx"
        Dim Nome_Aba_Arquivo As String = "Plan1"
        Dim Item_Not_Execute As String
        Rota = "33"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else

        End If
    End Sub

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_11()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12,1001"
        Dim vRows_Final As String = "1000,1107"
        Dim Unidade_Negocio_Origem As String = "L1489"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Ressuprimento_MG_DPSPMG.xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_MG"
        Dim Item_Not_Execute As String
        Rota = "02"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_12()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12,1001"
        Dim vRows_Final As String = "1000,1182"
        Dim Unidade_Negocio_Origem As String = "VD215"
        Dim Unidade_Negocio_Destino As String = "VD915"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Ressuprimento_BA.xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_BA"
        Dim Item_Not_Execute As String
        Rota = "02"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Ressuprimentos_Cenario_15()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6,1001"
        Dim vRows_Final As String = "1000,1101"
        Dim Unidade_Negocio_Origem As String = "VD945"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Ressuprimento_MG_DPSPMG (2).xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_MG"
        Dim Item_Not_Execute As String

        Rota = "08"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Transferencia_Cenario_20()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6,1001"
        Dim vRows_Final As String = "1000,1176"
        Dim Unidade_Negocio_Origem As String = "VD906"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\VD 909 . VD906 - CT20 - PR05 (8).xlsx"
        Dim Nome_Aba_Arquivo As String = "VD 909 > VD906 - CT20 - PR05"
        Dim Item_Not_Execute As String
        Rota = "906"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Transferencia_Cenario_22()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6,1001"
        Dim vRows_Final As String = "1000,1176"
        Dim Unidade_Negocio_Origem As String = "VD909"
        Dim Unidade_Negocio_Destino As String = "VD906"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\VD906 . VD909 - CT22 - PR 05_6213.xlsx"
        Dim Nome_Aba_Arquivo As String = "VD906 > VD909 - CT22 - PR05"
        Dim Item_Not_Execute As String
        Rota = "909"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else

        End If
    End Sub

    <TestMethod>
    Public Sub Transferencia_Cenario_25()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6,1001"
        Dim vRows_Final As String = "1000,1176"
        Dim Unidade_Negocio_Origem As String = "VD906"
        Dim Unidade_Negocio_Destino As String = "VD908"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\VD908 . VD906 - CT25 - PR05.xlsx"
        Dim Nome_Aba_Arquivo As String = "TODOS"
        Dim Item_Not_Execute As String
        Rota = "906"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else

        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_36()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6"
        Dim vRows_Final As String = "21"
        Dim Unidade_Negocio_Origem As String = "VD909"
        Dim Unidade_Negocio_Destino As String = "VD910"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Ressuprimento_MG_DPSPMG (1).xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_MG"
        Dim Item_Not_Execute As String
        Rota = "909"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Retorno = Script_BD(0, "PS_DPSP_HDR_DEVO_COLETOR_SP", Data_Retorno_Coletor, "", "", "").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                Mensagem = Script_BD(1, "PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(1, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(2, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_39()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12"
        Dim vRows_Final As String = "591"
        Dim Unidade_Negocio_Origem As String = "VD909"
        Dim Unidade_Negocio_Destino As String = "ST910"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Ressuprimento_MG_DPSPM_39.xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_MG"
        Dim Item_Not_Execute As String
        Rota = "909"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Retorno = Script_BD(0, "PS_DPSP_HDR_DEVO_COLETOR_SP", Data_Retorno_Coletor, "", "", "").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                Mensagem = Script_BD(1, "PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(1, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If

            If (Mensagem <> "") Then
                Mensagem = ""
                Mensagem = Script_BD(2, "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP", Retorno(Quantidade_Registro_Coletor - 1), Retorno(0), "", "")
            Else
                Debug.Fail("PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_42()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6,1001"
        Dim vRows_Final As String = "1000,1101"
        Dim Unidade_Negocio_Origem As String = "VD909"
        Dim Unidade_Negocio_Destino As String = "VD915"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Venda_MG_.xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_MG"
        Dim Item_Not_Execute As String
        Rota = "909"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_48()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12,1001"
        Dim vRows_Final As String = "1000,1182"
        Dim Unidade_Negocio_Origem As String = "VD915"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Venda_BA.xlsx"
        Dim Nome_Aba_Arquivo As String = "Ressuprimento_BA"
        Dim Item_Not_Execute As String
        Rota = "915"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_49()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12"
        Dim vRows_Final As String = "33"
        Dim Unidade_Negocio_Origem As String = "VD910"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Venda_SP - n49.xlsx"
        Dim Nome_Aba_Arquivo As String = "venda_sp"
        Dim Item_Not_Execute As String
        Rota = "910"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_50()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "6"
        Dim vRows_Final As String = "601"
        Dim Unidade_Negocio_Origem As String = "ST910"
        Dim Unidade_Negocio_Destino As String = "VD909"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Venda_SP - m.xlsx"
        Dim Nome_Aba_Arquivo As String = "venda_sp"
        Dim Item_Not_Execute As String
        Rota = "ST910"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else
            Mensagem = ""
            Retorno = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "ORDER_NO").Split(New Char() {","})
            Retorno_2 = Script_BD(0, "PS_IN_DEMAND_COLETOR_BH_MG", Unidade_Negocio_Destino, Data_Alteracao_Status, Unidade_Negocio_Origem, "PICK_BATCH_ID").Split(New Char() {","})
            Quantidade_Registro_Coletor = Retorno.Count() - 1

            If (Quantidade_Registro_Coletor > 0) Then
                For i = 0 To (Quantidade_Registro_Coletor - 1) Step 1
                    Mensagem = Script_BD(1, "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG", Unidade_Negocio_Destino, Retorno_2(i), Unidade_Negocio_Origem, Retorno(i))
                Next
            Else
                Debug.Fail("PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG")
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub Venda_Cenario_52()
        Dim URL_Cockpit_de_Ressuprimento, Click_Login, Status_Validacao, Quantidade_Registro, Quantidade_Rows, Rota As String
        Dim i, a As Int32
        Dim Conectado As Boolean
        Dim Lista_Itens As String()
        Dim Lista_Rows_Inicial As String()
        Dim Lista_Rows_Final As String()
        Dim Itens As String
        Dim vRows_Inicial As String = "12"
        Dim vRows_Final As String = "33"
        Dim Unidade_Negocio_Origem As String = "VD910"
        Dim Unidade_Negocio_Destino As String = "VD906"
        Dim Caminho_Massa_Dados As String = "C:\LeanTestAutomation\Massa de Dados DPSP\Venda_SP - n_52.xlsx"
        Dim Nome_Aba_Arquivo As String = "venda_sp"
        Dim Item_Not_Execute As String
        Rota = "910"

        Lista_Rows_Inicial = vRows_Inicial.Split(New Char() {","})
        Lista_Rows_Final = vRows_Final.Split(New Char() {","})

        Conectado = Coxexao_BD()
        If (Parte_Execucao = 1) Then
            If CBool(Conectado) Then
                Quantidade_Rows = vRows_Inicial.Split(New Char() {","}).Count - 1

                'Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                'Quantidade_Registro = Lista_Itens.Count - 1
                'For a = 0 To Quantidade_Registro Step 1
                '    Item_Not_Execute = Item_Not_Execute + Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Lista_Itens(a), "")
                'Next
                'Dim TESTE = Item_Not_Execute.Split(New Char() {","}).Count
                Script_BD(2, "ps_in_demand_Status", Unidade_Negocio_Destino, Data_Alteracao_Status, "", "")

                For i = 0 To Quantidade_Rows Step 1
                    Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(i), Lista_Rows_Final(i), Caminho_Massa_Dados, Nome_Aba_Arquivo)
                    Script_BD(3, "ps_dsp_interm_lote", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "ps_in_demand", Unidade_Negocio_Origem, Unidade_Negocio_Destino, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Origem, Itens, "0", "")
                    Script_BD(2, "PS_PHYSICAL_INV", Unidade_Negocio_Origem, Itens, "", "")


                    Script_BD(2, "PS_DEFAULT_LOC_INV", Unidade_Negocio_Destino, Itens, "", "")
                    Script_BD(2, "PS_BU_ITEMS_INV", Unidade_Negocio_Destino, Itens, "1000", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_HDR_DEVO.
                    Script_BD(2, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Itens, "", "") 'Fazendo o insert da ordem de compra na tabela PS_DPSP_LINE_DEVO.
                Next
                Lista_Itens = Resgatar_Valor_Coluna_Excel("A", Lista_Rows_Inicial(0), Lista_Rows_Final(Quantidade_Rows), Caminho_Massa_Dados, Nome_Aba_Arquivo).Split(New Char() {","})
                Quantidade_Registro = Lista_Itens.Count - 1
                For i = 0 To Quantidade_Registro Step 1
                    Mensagem = Script_BD(1, "PS_DSP_LOTCNTL_INV", Unidade_Negocio_Destino, Lista_Itens(i), "", "")
                    If (Mensagem <> "") Then
                        Exit For
                    End If
                Next
            End If
        Else

        End If
    End Sub

    Public Function Coxexao_BD() As Boolean
        Try
            Dim FDPSPHML = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP) (HOST=172.18.0.61)(PORT=1521)) (CONNECT_DATA=(SERVER=dedicated)(SERVICE_NAME=DBPSHML)));"

            Dim sConnString As String = "Provider=MSDAORA;Data Source=172.18.0.61:1521/DBPSHML;User Id=PS_FELIAS;Password=ps_5017;"
            Conexao = New OleDb.OleDbConnection(sConnString)
            'Test.Wait(10000)
            Conexao.Open()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function Script_BD(Action As Integer, Table_Name As String, Parameter_One As String, Parameter_Two As String, Parameter_Three As String, Parameter_Four As String) As String
        Try
            Dim Row_Execute As String
            Dim Sql_Query_Exec As OleDb.OleDbCommand
            Dim Reader As OleDb.OleDbDataReader
            Dim Returno_Coletor_SP As String

            If (Action = 0) Then 'Select
                Query_str = Select_Script(Table_Name, Parameter_One, Parameter_Two, Parameter_Three)
            ElseIf (Action = 1) Then 'Insert
                Query_str = Insert_Script(Table_Name, Parameter_One, Parameter_Two, Parameter_Three, Parameter_Four)
            ElseIf (Action = 2) Then 'Update
                Query_str = Update_Script(Table_Name, Parameter_One, Parameter_Two, Parameter_Three)
            ElseIf (Action = 3) Then 'Delete
                Query_str = Delete_Script(Table_Name, Parameter_One, Parameter_Two, Parameter_Three)
            End If

            If (Action <> 0) Then

                Sql_Query_Exec = New OleDb.OleDbCommand(Query_str, Conexao)

                Row_Execute = Sql_Query_Exec.ExecuteNonQuery()

                If (Row_Execute = 0) Then
                    Return Row_Execute.ToString() + ","
                Else
                    Return Row_Execute
                End If

            Else
                Sql_Query_Exec = New OleDb.OleDbCommand(Query_str, Conexao)

                Reader = Sql_Query_Exec.ExecuteReader()

                Dim i As Int32 = 0

                While Reader.Read()
                    Returno_Coletor_SP = Returno_Coletor_SP + Reader(Parameter_Four.ToString()).ToString() + ","
                End While
                If (Returno_Coletor_SP <> "") Then
                    Return Returno_Coletor_SP
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            Return ex.ToString()
            'Test.TestLog(Table_Name.ToString(), Table_Name.ToString(), Table_Name.ToString(), LibraryGlobal.LibGlobal.typelog.Failed)
        End Try
    End Function

    Public Function Close_DB()
        Conexao.Close()
    End Function

    Private Function Insert_Script(Table_Name As String, Parameter_One As String, Parameter_Two As String, Parameter_Three As String, Parameter_Four As String) As String
        If (Table_Name = "PS_DSP_LOTCNTL_INV") Then
            Query_str = "Insert into FDSPPRD.PS_DSP_LOTCNTL_INV " +
                             "(BUSINESS_UNIT, INV_ITEM_ID, DSP_LOTE_ID, DESCR, LOT_BIRTHDATE, " +
                             "EXPIRATION_DATE, DSP_LOT_STATUS, QTY_AVAILABLE, DT_TIMESTAMP, OPRID) " +
                             "Values " +
                             "('" + Parameter_One.ToString() + "', '" + Parameter_Two.ToString() + "', 'ABC', ' ', TO_DATE('01/01/2017 00:00:00', 'MM/DD/YYYY HH24:MI:SS'), " +
                             "TO_DATE('01/01/2020 00:00:00', 'MM/DD/YYYY HH24:MI:SS'), 'A', 1000, SYSDATE, 'CARGA')"

        ElseIf (Table_Name = "PS_DPSP_INDMD_POS_RETORNO_COLETOR_BH_MG") Then
            Query_str = "INSERT INTO fdspprd.PS_DPSP_INDMD_POS " +
                        "(SELECT business_unit, pick_batch_id, picklist_line_no, inv_item_id, '123', " +
                        " 'VOL1', qty_allocated, '01/01/2018', '01/12/2015', ' ', ' ', 0, " +
                        "'G', SYSDATE,' ' " +
                        "FROM fdspprd.ps_in_demand " +
                        "WHERE business_unit = '" + Parameter_One.ToString() + "' " +
                        "AND in_fulfill_state NOT IN ('70', '90', '30') " +
                        "And PICK_BATCH_ID >='" + Parameter_Two.ToString() + "' " +
                        "And TO_DATE(demand_date, 'YYYY-MM-DD') >= TO_DATE('" + Data_Alteracao_Status.ToString() + "', 'YYYY-MM-DD') " +
                        "AND ORDER_NO >='" + Parameter_Four.ToString() + "') "
            '"AND destin_bu = '" + Parameter_Three.ToString() + "' " +

        ElseIf (Table_Name = "PS_DPSP_HDR_RET_DV_RETORNO_COLETOR_SP") Then
            Query_str = "INSERT INTO FDSPPRD.PS_DPSP_HDR_RET_DV " +
                         "Select HD.DPSP_WMS_TP_INTERF As DPSP_WMS_TP_INTERF , " +
                         "TO_CHAR(SYSDATE ,'YYYYMMDDHH24') " +
                         "|| '00000000'         AS DPSP_WMS_TIMESTAMP , " +
                         " 'NEW'                 AS DPSP_WMS_STATUS , " +
                         "'INSERT'              AS DPSP_WMS_ACTION , " +
                         "HD.DPSP_WMS_ORDER_NO  As DPSP_WMS_ORDER_NO , " +
                         "HD.DPSP_WMS_PARTNR_NO As DPSP_WMS_PARTNR_NO , " +
                         "HD.DPSP_WMS_ROUTE     As DPSP_WMS_ROUTE , " +
                         "HD.DPPS_WMS_ORDER_P_N As DPPS_WMS_ORDER_P_N , " +
                         "' '                   AS DPSP_WMS_ERR_COD3 , " +
                         "' '                   AS DPSP_WMS_ERROR_COD , " +
                         "' '                   AS DPSP_WMS_INF_ERROR " +
                         "From FDSPPRD.PS_DPSP_HDR_DEVO HD " +
                         "Where HD.DPSP_WMS_TP_INTERF = 'CAD_ORD' " +
                         "And HD.DPSP_WMS_ORDER_NO >='" + Parameter_One.ToString() + "' " +
                         "And HD.DPSP_WMS_ORDER_NO <='" + Parameter_Two.ToString() + "' " +
                         "And Not EXISTS " +
                         "(SELECT 'X' " +
                         "From FDSPPRD.PS_DPSP_HDR_RET_DV HR " +
                         "Where HR.DPSP_WMS_ORDER_NO = HD.DPSP_WMS_ORDER_NO " +
                         "And HD.DPSP_WMS_TP_INTERF  = HD.DPSP_WMS_TP_INTERF)"

        ElseIf (Table_Name = "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP") Then
            Query_str = "INSERT INTO FDSPPRD.PS_DPSP_LIN_RET_DV " +
                        "Select LD.DPSP_WMS_TP_INTERF As DPSP_WMS_TP_INTERF , " +
                        "TO_CHAR(SYSDATE ,'YYYYMMDDHH24') " +
                        "|| '00000000'         AS DPSP_WMS_TIMESTAMP , " +
                        "'NEW'                 AS DPSP_WMS_STATUS , " +
                        "'INSERT'              AS DPSP_WMS_ACTION , " +
                        "LD.DPSP_WMS_ORDER_NO  AS DPSP_WMS_ORDER_NO , " +
                        "LD.DPSP_WMS_ORDER_POS AS DPSP_WMS_ORDER_POS , " +
                        "0                     As DPSP_WMS_ORD_POS_S , " +
                        "' '                   AS DPSP_WMS_USER_ID , " +
                        "' '                   AS DPSP_WMS_SO_ID , " +
                        "LD.DPSP_WMS_ARTICLE_N AS DPSP_WMS_ARTICLE_N , " +
                        "LD.DPSP_WMS_STOCK_TYP AS DPSP_WMS_STOCK_TYP , " +
                        "LD.DPSP_WMS_SU1_DESC  AS DPSP_WMS_SU1_DESC , " +
                        "LD.dpsp_wms_su1_order AS DPSP_WMS_SU1_DELVD , " +
                        " ' '                   AS DPSP_WMS_CANCELLAT , " +
                        "NULL                  AS EXPIRATION_DATE , " +
                        "LD.DPSP_WMS_LOTE      AS DPSP_WMS_LOTE , " +
                        "' '                   AS DPSP_WMS_INFO , " +
                        "' '                   AS DPSP_WMS_ERROR_COD , " +
                        "' '                   AS DPSP_WMS_INF_ERROR " +
                        "FROM FDSPPRD.PS_DPSP_LINE_DEVO LD " +
                        "WHERE LD.DPSP_WMS_TP_INTERF = 'CAD_ORD' " +
                        "And LD.DPSP_WMS_ORDER_NO   >='" + Parameter_One.ToString() + "' " +
                        "AND LD.DPSP_WMS_ORDER_NO   <='" + Parameter_Two.ToString() + "' " +
                        "And Not EXISTS " +
                        "(SELECT 'X' " +
                        "FROM FDSPPRD.PS_DPSP_LIN_RET_DV LR " +
                        "WHERE LR.DPSP_WMS_ORDER_NO = LD.DPSP_WMS_ORDER_NO " +
                        "And LR.DPSP_WMS_TP_INTERF  = LD.DPSP_WMS_TP_INTERF)"

        Else
            'Test.TestLog("Table Name:  " + Table_Name, "Table Name: " + Table_Name, "Table Name: " + Table_Name, typelog.Failed)
        End If
        If (Query_str <> "") Then
            Return Query_str
        Else
            Return False
        End If
    End Function

    Private Function Update_Script(Table_Name As String, Parameter_One As String, Parameter_Two As String, Parameter_Three As String) As String
        If (Table_Name = "ps_in_demand") Then
            Query_str = "UPDATE fdspprd.PS_IN_DEMAND SET IN_FULFILL_STATE = '90' " +
                            "WHERE BUSINESS_UNIT = '" + Parameter_One.ToString() + "' " +
                            "AND DESTIN_BU = '" + Parameter_Two.ToString() + "' " +
                            "And IN_FULFILL_STATE Not IN ('70','90')"

        ElseIf (Table_Name = "PS_BU_ITEMS_INV") Then
            If (Parameter_Three = 0) Then
                Query_str = "UPDATE fdspprd.PS_BU_ITEMS_INV Set QTY_AVAILABLE = 0, QTY_RESERVED = 0, QTY_OWNED = 0, QTY_ONHAND = 0 " +
                            "WHERE BUSINESS_UNIT = '" + Parameter_One.ToString() + "' AND INV_ITEM_ID IN (" + Parameter_Two.ToString() + ") "
            ElseIf (Parameter_Three = "1000") Then
                Query_str = "UPDATE fdspprd.PS_BU_ITEMS_INV SET QTY_AVAILABLE = QTY_AVAILABLE + 3000, QTY_OWNED = QTY_OWNED + 3000, QTY_ONHAND = QTY_ONHAND + 3000 " +
                            "WHERE BUSINESS_UNIT = '" + Parameter_One.ToString() + "' AND INV_ITEM_ID IN (" + Parameter_Two.ToString() + ") "
            End If

        ElseIf (Table_Name = "PS_PHYSICAL_INV") Then
            Query_str = "UPDATE fdspprd.PS_PHYSICAL_INV SET QTY = 0, QTY_RESERVED = 0, QTY_BASE = 0, QTY_RESERVED_BASE = 0 " +
                            "WHERE BUSINESS_UNIT = '" + Parameter_One.ToString() + "' AND INV_ITEM_ID IN (" + Parameter_Two.ToString() + ") "

        ElseIf (Table_Name = "PS_DEFAULT_LOC_INV") Then
            Query_str = "UPDATE fdspprd.PS_PHYSICAL_INV A SET A.QTY = A.QTY + 3000, A.QTY_BASE = A.QTY_BASE + 3000 " +
                            "WHERE (A.BUSINESS_UNIT, A.INV_ITEM_ID, A.STORAGE_AREA, A.STOR_LEVEL_1, A.STOR_LEVEL_2, A.STOR_LEVEL_3, A.STOR_LEVEL_4) In " +
                            "(SELECT B.BUSINESS_UNIT, B.INV_ITEM_ID, B.STORAGE_AREA, B.STOR_LEVEL_1, B.STOR_LEVEL_2, B.STOR_LEVEL_3, B.STOR_LEVEL_4  " +
                            "FROM fdspprd.PS_DEFAULT_LOC_INV B " +
                            "WHERE B.BUSINESS_UNIT = '" + Parameter_One.ToString() + "' " +
                            "And B.INV_ITEM_ID IN (" + Parameter_Two.ToString() + ") " +
                            "And B.DEF_LOC_TYPE = 'O') "

        ElseIf (Table_Name = "PS_DSP_LOTCNTL_INV") Then
            Query_str = "UPDATE fdspprd.PS_DSP_LOTCNTL_INV A SET A.QTY_AVAILABLE = QTY_AVAILABLE + 1000 " +
                            "WHERE A.BUSINESS_UNIT = '" + Parameter_One.ToString() + "' " +
                            "AND A.INV_ITEM_ID IN (" + Parameter_Two.ToString() + ") " +
                            "And A.DSP_LOTE_ID = (SELECT MAX(B.DSP_LOTE_ID) FROM  fdspprd.PS_DSP_LOTCNTL_INV B " +
                            "WHERE B.BUSINESS_UNIT = A.BUSINESS_UNIT " +
                            "And B.INV_ITEM_ID = A.INV_ITEM_ID)"

        ElseIf (Table_Name = "ps_in_demand_Status") Then
            Query_str = "UPDATE fdspprd.ps_in_demand " +
                         "Set IN_FULFILL_STATE = '90' " +
                         "WHERE BUSINESS_UNIT = ' " + Parameter_One.ToString() + "' " +
                         "And IN_FULFILL_STATE Not In ( '90','70','95') " +
                         "AND DEMAND_DATE <= '" + Parameter_Two.ToString() + "'"

        ElseIf (Table_Name = "PS_DPSP_LIN_RET_DV_RETORNO_COLETOR_SP") Then
            Query_str = "UPDATE FDSPPRD.PS_DPSP_LIN_RET_DV " +
                        "Set DPSP_WMS_LOTE        = 'ABC',  " +
                        "EXPIRATION_DATE        = TO_DATE('2019-01-01', 'YYYY-MM-DD')  " +
                        "WHERE DPSP_WMS_TP_INTERF = 'CAD_ORD'  " +
                        "AND DPSP_WMS_ORDER_NO   >='" + Parameter_One.ToString() + "'  " +
                        "AND DPSP_WMS_ORDER_NO   <='" + Parameter_Two.ToString() + "'"
        Else
            'Test.TestLog("Table Name: " + Table_Name, "Table Name: " + Table_Name, "Table Name: " + Table_Name, typelog.Failed)
        End If
        If (Query_str <> "") Then
            Return Query_str
        Else
            Return False
        End If
    End Function

    Private Function Select_Script(Table_Name As String, Parameter_One As String, Parameter_Two As String, Parameter_Three As String) As String
        If (Table_Name = "PS_DPSP_HDR_DEVO") Then
            Query_str = "SELECT HD.DPSP_WMS_TP_INTERF AS DPSP_WMS_TP_INTERF " +
                        ", TO_CHAR(SYSDATE " +
                        ",'YYYYMMDDHH24') || '00000000' AS DPSP_WMS_TIMESTAMP " +
                        ", 'NEW' AS DPSP_WMS_STATUS " +
                        ", 'INSERT' AS DPSP_WMS_ACTION " +
                        ", HD.DPSP_WMS_ORDER_NO AS DPSP_WMS_ORDER_NO " +
                        " HD.DPSP_WMS_PARTNR_NO AS DPSP_WMS_PARTNR_NO " +
                        ", HD.DPSP_WMS_ROUTE AS DPSP_WMS_ROUTE " +
                        ", HD.DPPS_WMS_ORDER_P_N AS DPPS_WMS_ORDER_P_N " +
                        ", ' ' AS DPSP_WMS_ERR_COD3 " +
                        ", ' ' AS DPSP_WMS_ERROR_COD " +
                        ", ' ' AS DPSP_WMS_INF_ERROR " +
                        "FROM FDSPPRD.PS_DPSP_HDR_DEVO HD " +
                        "WHERE HD.DPSP_WMS_TP_INTERF = 'CAD_ORD'" +
                        "AND NOT EXISTS ( " +
                        "SELECT 'X' " +
                        "FROM FDSPPRD.PS_DPSP_HDR_RET_DV HR " +
                        "WHERE HR.DPSP_WMS_ORDER_NO = HD.DPSP_WMS_ORDER_NO " +
                        "AND HD.DPSP_WMS_TP_INTERF = HD.DPSP_WMS_TP_INTERF)"

        ElseIf (Table_Name = "PS_DPSP_HDR_DEVO_COLETOR_SP") Then
            Query_str = "SELECT HD.DPSP_WMS_ORDER_NO AS DPSP_WMS_ORDER_NO " +
                        "FROM FDSPPRD.PS_DPSP_HDR_DEVO HD " +
                        "WHERE HD.DPSP_WMS_TP_INTERF = 'CAD_ORD' " +
                        "And  HD.DELIVERY_DT >= TO_DATE('" + Parameter_One.ToString() + "', 'YYYY-MM-DD')" +
                        "AND NOT EXISTS ( " +
                        "Select 'X' " +
                        "FROM FDSPPRD.PS_DPSP_HDR_RET_DV HR " +
                        "WHERE HR.DPSP_WMS_ORDER_NO = HD.DPSP_WMS_ORDER_NO " +
                        "AND HD.DPSP_WMS_TP_INTERF = HD.DPSP_WMS_TP_INTERF) " +
                        "ORDER BY DPSP_WMS_ORDER_NO DESC"

        ElseIf (Table_Name = "PS_IN_DEMAND_COLETOR_BH_MG") Then
            Query_str = "SELECT DISTINCT ORDER_NO  , PICK_BATCH_ID, PICK_DTTM " +
                        "FROM FDSPPRD.PS_IN_DEMAND " +
                        "WHERE BUSINESS_UNIT = '" + Parameter_One.ToString() + "' " +
                        "And DEMAND_DATE >= TO_DATE('" + Parameter_Two.ToString() + "', 'YYYY-MM-DD') " +
                        "AND IN_FULFILL_STATE NOT IN ('300','70','900','95')" +
                        "ORDER BY ORDER_NO ASC"

            '"AND DESTIN_BU = '" + Parameter_Three.ToString() + "' " +
            '"--And ORDER_NO >= 'PED6408183' " +
            '"--AND ROUTE_CD = '10' " +
            '"--And DEMAND_SOURCE Not IN ('PL') " +      
        Else
            'Test.TestLog("Table Name: " + Table_Name, "Table Name: " + Table_Name, "Table Name: " + Table_Name, typelog.Failed)
        End If
        If (Query_str <> "") Then
            Return Query_str
        Else
            Return False
        End If
    End Function

    Private Function Delete_Script(Table_Name As String, Parameter_One As String, Parameter_Two As String, Parameter_Three As String) As String
        If (Table_Name = "ps_dsp_interm_lote") Then
            Query_str = "DELETE FROM fdspprd.PS_DSP_INTERM_LOTE " +
                            "WHERE DSP_PROCESSO = 'RE' " +
                            "AND (BUSINESS_UNIT_IN, DEMAND_SOURCE, BUS_UNIT_SOURCE, ORDER_NO, ORDER_INT_LINE_NO, SCHED_LINE_NBR, INV_ITEM_ID, DEMAND_LINE_NO) IN " +
                            "(SELECT BUSINESS_UNIT, DEMAND_SOURCE, SOURCE_BUS_UNIT, ORDER_NO, ORDER_INT_LINE_NO, SCHED_LINE_NBR, INV_ITEM_ID, DEMAND_LINE_NO FROM fdspprd.PS_IN_DEMAND " +
                            "WHERE  BUSINESS_UNIT = '" + Parameter_One.ToString() + "' " +
                            "AND DESTIN_BU = '" + Parameter_Two.ToString() + "' " +
                            "And IN_FULFILL_STATE Not IN ('70','90'))"
        Else
            'Test.TestLog("Table Name: " + Table_Name, "Table Name: " + Table_Name, "Table Name: " + Table_Name, typelog.Failed)
        End If
        If (Query_str <> "") Then
            Return Query_str
        Else
            Return False
        End If
    End Function

    Function Resgatar_Valor_Coluna_Excel(Col As String, Row_Initial As Int32, Row_End As Int32, pathFile As String, Nome_Aba As String) As String
        Dim i, Quantidade, Row As Int32
        Dim Valor As String
        Try

            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet

            Quantidade = Row_End - (Row_Initial - 1)
            Row = Row_Initial
            xlWorkBook = xlApp.Workbooks.Open(pathFile)
            xlWorkSheet = xlWorkBook.Worksheets(Nome_Aba)
            For i = 0 To Quantidade - 1 Step 1
                If (i = 0) Then
                    Valor = xlWorkSheet.Cells(Row, Col).Value.ToString()
                    If (Quantidade > 1) Then
                        Valor = Valor.Replace(" ", "") + ","
                    Else
                        Valor = Valor.Replace(" ", "")
                    End If
                Else
                    Valor = Valor + xlWorkSheet.Cells((i + Row), Col).Value.ToString().Replace(" ", "")
                    If (i < (Quantidade - 1)) Then Valor = Valor + ","
                End If
            Next
            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()
            'caso queira abrir o excel

            Return Valor
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

End Class
