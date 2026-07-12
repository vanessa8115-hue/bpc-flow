# ==========================================
# BPC FLOW v1.4
# Controle de Beneficiários BPC
# Pesquisa e Relatórios
# ==========================================

import os
import json
from datetime import datetime, timedelta
from openpyxl import Workbook


ARQUIVO = "beneficiarios.json"
ARQUIVO_EXCEL = "BPC_Beneficiarios.xlsx"


beneficiarios = []


# ==========================================
# DADOS
# ==========================================

def carregar_dados():

    global beneficiarios

    try:

        with open(
            ARQUIVO,
            "r",
            encoding="utf-8"
        ) as arquivo:

            beneficiarios = json.load(arquivo)

    except FileNotFoundError:

        beneficiarios = []



def salvar_dados():

    with open(
        ARQUIVO,
        "w",
        encoding="utf-8"
    ) as arquivo:

        json.dump(
            beneficiarios,
            arquivo,
            indent=4,
            ensure_ascii=False
        )



# ==========================================
# TELA
# ==========================================

def limpar_tela():

    os.system(
        "cls" if os.name == "nt" else "clear"
    )


def linha():

    print("=" * 55)



def titulo(texto):

    linha()

    print(
        texto.center(55)
    )

    linha()



def pausar():

    input(
        "\nPressione ENTER para continuar..."
    )



# ==========================================
# ANALISAR CADASTRO
# ==========================================

def analisar_cadastro(data_atualizacao):

    try:

        data = datetime.strptime(
            data_atualizacao,
            "%d/%m/%Y"
        )

        vencimento = (
            data + timedelta(days=730)
        )

        hoje = datetime.now()


        if vencimento < hoje:

            situacao = "Cadastro vencido"


        else:

            dias = (
                vencimento - hoje
            ).days


            if dias <= 90:

                situacao = (
                    "Vence nos próximos 90 dias"
                )

            else:

                situacao = (
                    "Cadastro atualizado"
                )


        return (
            vencimento.strftime("%d/%m/%Y"),
            situacao
        )


    except:

        return (
            "Data inválida",
            "Não analisado"
        )



# ==========================================
# EXPORTAR EXCEL
# ==========================================

def exportar_excel():

    if not beneficiarios:

        print(
            "\nNenhum dado para exportar."
        )

        pausar()

        return


    planilha = Workbook()

    aba = planilha.active

    aba.title = "Beneficiários"


    aba.append([
        "Nome",
        "CPF",
        "NIS",
        "Benefício",
        "Localidade",
        "Ilha",
        "Grupo",
        "Atualização",
        "Vencimento",
        "Situação"
    ])


    for b in beneficiarios:

        aba.append([

            b.get("Nome"),
            b.get("CPF"),
            b.get("NIS"),
            b.get("Beneficio"),
            b.get("Localidade"),
            b.get("Ilha"),
            b.get("Grupo"),
            b.get("Ultima Atualizacao"),
            b.get("Vencimento Cadastro"),
            b.get("Situacao")

        ])


    planilha.save(
        ARQUIVO_EXCEL
    )


    print(
        "\n✅ Excel atualizado!"
    )

    print(
        "Arquivo:",
        ARQUIVO_EXCEL
    )


    pausar()


# ==========================================
# CADASTRO DE BENEFICIÁRIO
# ==========================================

def cadastrar_beneficiario():

    limpar_tela()
    titulo("CADASTRO DE BENEFICIÁRIO")


    b = {}


    b["Nome"] = input("Nome completo: ").strip()

    b["CPF"] = input("CPF: ").strip()

    b["NIS"] = input("NIS: ").strip()

    b["Nascimento"] = input("Nascimento: ").strip()

    b["Telefone"] = input("Telefone: ").strip()



    print("\nTipo de benefício:")
    print("1 - BPC Idoso")
    print("2 - BPC Pessoa com Deficiência")


    opcao = input("Escolha: ")


    if opcao == "1":

        b["Beneficio"] = "BPC Idoso"


    elif opcao == "2":

        b["Beneficio"] = "BPC Pessoa com Deficiência"


    else:

        b["Beneficio"] = "Não informado"




    print("\nLocalidade:")
    print("1 - Itacuruçá")
    print("2 - Ilha")


    opcao = input("Escolha: ")


    if opcao == "1":

        b["Localidade"] = "Itacuruçá"

        b["Ilha"] = "Não se aplica"



    elif opcao == "2":

        b["Localidade"] = "Ilha"


        print("\nEscolha a ilha:")

        print("1 - Ilha da Marambaia")

        print("2 - Ilha de Jaguanum")

        print("3 - Ilha de Itacuruçá")

        print("4 - Outra")


        ilha = input("Escolha: ")



        if ilha == "1":

            b["Ilha"] = "Ilha da Marambaia"


        elif ilha == "2":

            b["Ilha"] = "Ilha de Jaguanum"


        elif ilha == "3":

            b["Ilha"] = "Ilha de Itacuruçá"


        elif ilha == "4":

            b["Ilha"] = input(
                "Nome da ilha: "
            )


        else:

            b["Ilha"] = "Não informado"



    print("\nParticipação familiar:")

    print("1 - Responsável Familiar (RF)")

    print("2 - Componente do Grupo Familiar")


    grupo = input("Escolha: ")



    if grupo == "1":

        b["Grupo"] = "Responsável Familiar (RF)"


    elif grupo == "2":

        b["Grupo"] = "Componente do Grupo Familiar"


    else:

        b["Grupo"] = "Não informado"



    b["Ultima Atualizacao"] = input(
        "\nÚltima atualização (DD/MM/AAAA): "
    )



    vencimento, situacao = analisar_cadastro(
        b["Ultima Atualizacao"]
    )


    b["Vencimento Cadastro"] = vencimento

    b["Situacao"] = situacao



    beneficiarios.append(b)

    salvar_dados()



    print("\n✅ Cadastro realizado!")

    print(
        "Situação:",
        situacao
    )


    pausar()




# ==========================================
# PESQUISA POR NOME OU CPF
# ==========================================

def pesquisar_beneficiario():

    limpar_tela()

    titulo(
        "PESQUISAR BENEFICIÁRIO"
    )


    termo = input(
        "Digite nome ou CPF: "
    ).lower().strip()


    encontrou = False



    for b in beneficiarios:


        nome = b["Nome"].lower()

        cpf = b["CPF"].lower()



        if termo in nome or termo in cpf:


            encontrou = True


            linha()

            print(
                "Nome:",
                b["Nome"]
            )

            print(
                "CPF:",
                b["CPF"]
            )

            print(
                "Benefício:",
                b["Beneficio"]
            )

            print(
                "Local:",
                b["Localidade"]
            )

            print(
                "Ilha:",
                b["Ilha"]
            )

            print(
                "Situação:",
                b["Situacao"]
            )


            linha()



    if not encontrou:

        print(
            "\nNenhum resultado encontrado."
        )


    pausar()


# ==========================================
# LISTAR BENEFICIÁRIOS
# ==========================================

def listar_beneficiarios():

    limpar_tela()
    titulo("LISTA DE BENEFICIÁRIOS")


    if not beneficiarios:

        print("Nenhum beneficiário cadastrado.")


    else:

        for i, b in enumerate(beneficiarios, start=1):

            print(f"\n{i} - {b['Nome']}")
            print("CPF:", b["CPF"])
            print("Benefício:", b["Beneficio"])
            print("Localidade:", b["Localidade"])
            print("Ilha:", b["Ilha"])
            print("Situação:", b["Situacao"])

            linha()


    pausar()




# ==========================================
# PENDÊNCIAS
# ==========================================

def ver_pendencias():

    limpar_tela()
    titulo("CADASTROS PARA ATUALIZAÇÃO")


    encontrou = False


    for b in beneficiarios:


        if b["Situacao"] != "Cadastro atualizado":


            encontrou = True


            print("\nNome:", b["Nome"])
            print("CPF:", b["CPF"])
            print("Localidade:", b["Localidade"])
            print("Ilha:", b["Ilha"])
            print("Situação:", b["Situacao"])
            print("Vencimento:", b["Vencimento Cadastro"])

            linha()



    if not encontrou:

        print("Nenhuma pendência encontrada.")


    pausar()




# ==========================================
# RELATÓRIO POR TERRITÓRIO
# ==========================================

def relatorio_territorio():

    limpar_tela()
    titulo("RELATÓRIO TERRITORIAL")


    print("Escolha o território:")

    print("1 - Itacuruçá")

    print("2 - Ilha da Marambaia")

    print("3 - Ilha de Jaguanum")

    print("4 - Ilha de Itacuruçá")


    opcao = input("Escolha: ")



    filtro = ""


    if opcao == "1":

        filtro = "Itacuruçá"


    elif opcao == "2":

        filtro = "Ilha da Marambaia"


    elif opcao == "3":

        filtro = "Ilha de Jaguanum"


    elif opcao == "4":

        filtro = "Ilha de Itacuruçá"


    else:

        print("Opção inválida.")

        pausar()

        return



    total = 0

    idosos = 0

    pcd = 0

    atualizados = 0

    pendentes = 0



    for b in beneficiarios:


        local = b["Localidade"]

        ilha = b["Ilha"]



        if filtro == local or filtro == ilha:


            total += 1


            if b["Beneficio"] == "BPC Idoso":

                idosos += 1


            if b["Beneficio"] == "BPC Pessoa com Deficiência":

                pcd += 1


            if b["Situacao"] == "Cadastro atualizado":

                atualizados += 1

            else:

                pendentes += 1




    linha()

    print("Território:", filtro)

    print("Total:", total)

    print("BPC Idoso:", idosos)

    print("BPC PCD:", pcd)

    print("Atualizados:", atualizados)

    print("Pendentes:", pendentes)

    linha()


    pausar()




# ==========================================
# ESTATÍSTICAS
# ==========================================

def estatisticas():

    limpar_tela()
    titulo("ESTATÍSTICAS")


    total = len(beneficiarios)

    atualizados = 0

    pendentes = 0

    proximos = 0



    for b in beneficiarios:


        if b["Situacao"] == "Cadastro atualizado":

            atualizados += 1


        else:

            pendentes += 1



        if b["Situacao"] == "Vence nos próximos 90 dias":

            proximos += 1




    percentual = 0


    if total > 0:

        percentual = (
            atualizados / total
        ) * 100



    print("Total de beneficiários:", total)

    print()

    print(
        "Atualizados:",
        atualizados
    )

    print(
        "Pendentes:",
        pendentes
    )

    print(
        "Vencem em até 90 dias:",
        proximos
    )

    print()

    print(
        f"Percentual atualizado: {percentual:.1f}%"
    )


    pausar()


# ==========================================
# PAINEL DE ACOMPANHAMENTO
# ==========================================

def painel_acompanhamento():

    limpar_tela()

    titulo("PAINEL DE ACOMPANHAMENTO")


    total = len(beneficiarios)


    atualizados = 0
    vencendo = 0
    vencidos = 0


    idosos = 0
    pcd = 0


    itacuruca = 0
    ilhas = 0


    marambaia = 0
    jaguanum = 0
    ilha_itacuruca = 0



    for b in beneficiarios:


        # Situação do cadastro

        if b["Situacao"] == "Cadastro atualizado":

            atualizados += 1


        elif b["Situacao"] == "Vence nos próximos 90 dias":

            vencendo += 1


        elif b["Situacao"] == "Cadastro vencido":

            vencidos += 1



        # Tipo de benefício

        if b["Beneficio"] == "BPC Idoso":

            idosos += 1


        elif b["Beneficio"] == "BPC Pessoa com Deficiência":

            pcd += 1



        # Localidade

        if b["Localidade"] == "Itacuruçá":

            itacuruca += 1


        elif b["Localidade"] == "Ilha":

            ilhas += 1



        # Ilhas

        if b["Ilha"] == "Ilha da Marambaia":

            marambaia += 1


        elif b["Ilha"] == "Ilha de Jaguanum":

            jaguanum += 1


        elif b["Ilha"] == "Ilha de Itacuruçá":

            ilha_itacuruca += 1




    print("Total de beneficiários:", total)

    print()


    print("SITUAÇÃO DOS CADASTROS")

    print("🟢 Atualizados:", atualizados)

    print("🟡 Vencem em até 90 dias:", vencendo)

    print("🔴 Vencidos:", vencidos)


    print()

    print("TIPO DE BENEFÍCIO")

    print("BPC Idoso:", idosos)

    print(
        "BPC Pessoa com Deficiência:",
        pcd
    )


    print()

    print("LOCALIDADE")

    print("Itacuruçá:", itacuruca)

    print("Ilhas:", ilhas)


    print()

    print("ILHAS")

    print("Marambaia:", marambaia)

    print("Jaguanum:", jaguanum)

    print(
        "Ilha de Itacuruçá:",
        ilha_itacuruca
    )


    linha()

    pausar()

# ==========================================
# MENU PRINCIPAL
# ==========================================

def menu():

    while True:


        limpar_tela()

        titulo("BPC FLOW v1.4")


        print("1 - Painel de acompanhamento")

        print("2 - Cadastrar beneficiário")

        print("3 - Consultar por CPF")

        print("4 - Pesquisar beneficiário")

        print("5 - Listar beneficiários")

        print("6 - Ver pendências")

        print("7 - Estatísticas")

        print("8 - Relatório territorial")

        print("9 - Exportar para Excel")

        print("0 - Sair")


        opcao = input(
            "Escolha uma opção: "
        )



        if opcao == "1":

            painel_acompanhamento()

        elif opcao == "2":    

            cadastrar_beneficiario()


        elif opcao == "3":

            pesquisar_beneficiario()


        elif opcao == "4":

            pesquisar_beneficiario()


        elif opcao == "5":

            listar_beneficiarios()


        elif opcao == "6":

            ver_pendencias()


        elif opcao == "7":

            estatisticas()


        elif opcao == "8":

            relatorio_territorio()


        elif opcao == "9":

            exportar_excel()


        elif opcao == "0":

            print("\nEncerrando BPC FLOW...")

            break


        else:

            print("\nOpção inválida!")

            pausar()




# ==========================================
# INÍCIO
# ==========================================

carregar_dados()


if __name__ == "__main__":


    menu()
    
           