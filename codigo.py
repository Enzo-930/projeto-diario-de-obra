# importando funcões
import os
from openpyxl import load_workbook

#Função limpar tela

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')


# perguntas 
def questoes ():
    data = input("\nQual a data de hoje?(escreva no modelo exemplo: 24 de abril) ")
    cnd_clim = input ("\nComo está o clima? ")
    funcionarios = input ("\nQuantos funcionários vieram hoje? ")
    while True:
        ocorrencia = input ("\nHouve alguma ocorrência?(sim/não) ").strip() .lower()
        if ocorrencia == "sim":
            q_ocorrencia = input ("\nQuais ocorrências? ")
            break
        elif ocorrencia == "não":
            break
        else:
            print ("\nResposta inválida.Tente novamente")

    while True:
        impedimento = input ("\nHouve condições de impedimento?(sim/não) ") .strip() .lower()
        if impedimento == "sim":
            q_condicoes = input ("\nQual o motivo? " )
            break
        elif impedimento == "não":
            break
        else:
            print("\nResposta inválida. Tente novamente.")

    atv_exec = input ("\nQual atividade foi executada? ")
    while True:
        mat_rec = input ("\nAlgum material recebido?(sim/não) ").strip () .lower()
        if mat_rec == "sim":
            q_materiais = input ("\nQuais materiais? " )
            break
        elif mat_rec == "não":
            break
        else: 
            print("\nReposta Inválida. Tente novamente.")

 #interação com respostas
    limpar_tela ()
    print(f"\nA data do dia é: {data}")
    print(f"\nO clima está: {cnd_clim}")
    print(f"\nVieram {funcionarios} funcionarios")
    print(f"\n{ocorrencia} houveram ocorrências")
    if ocorrencia == "sim":
        print(f"sendo suas elas: {q_ocorrencia}")
    else:
        0
    print(f"\n{impedimento} ocorreu impedimento do trabalho")
    if impedimento == "sim":
        print(f"{q_condicoes} ")
    else: 
        0
    print(f"\nforam executadas as atividades:{atv_exec}")
    print(f"\n{mat_rec} foram recebidos materiais")
    if mat_rec == "sim":
        print(f"entre eles:{q_materiais}")
    else:
        0
    while True:
        resp_edit = input("\n\n\n\n\n\n\n\n\n\n\nDeseja editar alguma resposta?(s/n)").strip().lower()
        if resp_edit == "n":
            break
        
        else:
            limpar_tela ()
            print("\nData [1]")
            print("\nClima [2]")
            print("\nFuncionários [3]")
            print("\nOcorrências (se houve ou não)[4]")
            if ocorrencia == "sim":
                print("\nQual ocorrência [4.5]")
            else:
                0
            print("\nImpedimento do trabalho [5]")
            if impedimento == "sim":
                print("\nQual foi o impedimento [5.5]")
            else:
                0
            print("\nAtividades executadas [6]")
            print("\nMateriass [7]")
            if mat_rec == "sim":
                print("\nQuais materiais [7.5]")
            else:
                0
            while True:
                oq_editar = input("Digite o número correspondete que deseja editar: ")
                if oq_editar == "1":
                    data = input("\nDigite sua data correta: ")
                    break
                
                elif oq_editar == "2":
                    cnd_clim = input("\nO clima correto: ")
                    break
                
                elif oq_editar == "3":
                    funcionarios = input("\nQuantos funcionários vieram? ")
                    break
                
                elif oq_editar == "4":
                    while True:
                        ocorrencia = input("\nHouve alguma ocorrência?(sim/não) ")
                        if ocorrencia == "sim":
                            q_ocorrencia = input ("\nQuais ocorrências? " )
                            break
                        elif ocorrencia == "não":
                            q_ocorrencia = 0
                            break
                        else:
                            print ("\nResposta inválida, tente novamente.")
                    break   
                
                elif oq_editar == "4.5":
                    q_ocorrencia = input ("\nQuais ocorrências? ")
                    break

                elif oq_editar == "5":
                    while  True:
                        impedimento = input ("\nHouve impedimento do trabalho? ")
                        if impedimento == "sim":
                            q_condicoes = input ("\nQual a causa? ")
                            break
                        elif impedimento == "não":
                            q_condicoes = 0
                            break
                        else:
                            print("\nResposta inválida, tente novamente. ")
                    break
                            
                elif oq_editar == "5.5":
                    q_condicoes = input("\nQual o motivo de impedimênto? ")
                    break

                elif oq_editar == "6":
                    atv_exec = input("\nQual atividade foi executada? ")
                    break
                
                elif oq_editar == "7":
                    while True:
                        mat_rec = input ("\nHouve algum material recebido? ")
                        if mat_rec == "sim":
                            q_materiais = input ("Quais materiais? ")
                            break
                        elif mat_rec == "não":
                            q_materiais = 0
                            break
                        else:
                            print("\nResposta inválida, tente novamente. ")
                    break
                
                elif oq_editar == "7.5":
                    q_materiais = input("\nQuais materiais foram recebidos? ")
                    break
                
                else:
                    print("\n\nNão foi escolhido nenhum número válido, digite novamente. ")
    limpar_tela ()
    print("\n    ---- Suas informações do dia foram salvas, te espero amanhã 👍----")        

                        
                        

 # edição do documento


    arquivo = load_workbook("diario de obra.xlsx")
    aba_perg = arquivo["Planilha1"]

    aba_perg["A5"].value = f"{data}"
    aba_perg["A9"].value = f"{cnd_clim}"
    aba_perg["A13"].value = f"{impedimento}"
    if impedimento == "sim":
        aba_perg["A14"].value = f"{q_condicoes}"
    else:
        aba_perg["A14"].value = "Não houve impedimento"
    aba_perg["A19"].value = f"{ocorrencia}"
    if ocorrencia == "sim":
        aba_perg["A20"].value = f"{q_ocorrencia}"
    else:
        aba_perg["A20"].value = "Não houve nenhuma ocorrência."
    aba_perg["A24"].value = f"{funcionarios}"
    aba_perg["A28"].value = f"{atv_exec}"
    aba_perg["A33"].value = f"{mat_rec}"
    if mat_rec == "sim":
        aba_perg["A34"].value = f"{q_materiais}"
    else:
        aba_perg["A34"].value = "Não houveram materiais recebidos hoje."


    arquivo.save(f"diario de obra {data}.xlsx")

#executa 
questoes ()

