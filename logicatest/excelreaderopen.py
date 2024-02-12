import openpyxl as open

# Carregar o arquivo Excel
planilha = open.load_workbook("CópiadeEngenhariadeSoftware-DesafioCauêTamiarana.xlsx")

# Acessar a aba ativa
aba_ativa = planilha.active

# Número de aulas e porcentagem máxima de faltas
tot_aulas = 60
max_falt = 25

# Iterar sobre as linhas da planilha (a partir da linha 4)
for linha in aba_ativa.iter_rows(min_row=4):
    
    # Acessar os valores de cada coluna
    faltas = linha[2].value  # Índice 2 corresponde à coluna C (faltas)
    p1 = linha[3].value      # Índice 3 corresponde à coluna D (P1)
    p2 = linha[4].value      # Índice 4 corresponde à coluna E (P2)
    p3 = linha[5].value      # Índice 5 corresponde à coluna F (P3)
    situacao = linha[6]      # Índice 6 corresponde à coluna G
    nota_final = linha[7]    # Índice 7 corresponde à coluna H

    # Verificar a situação do aluno e calcular a nota final
    if (faltas / tot_aulas) * 100 > max_falt:
        situacao.value = "Reprovado por faltas"
        nota_final.value = 0
    else:
        media = (p1 + p2 + p3) / 3
        if media < 50:
            situacao.value = "Reprovado por nota"
            nota_final.value = 0
        elif 50 <= media < 70:
            situacao.value = "Exame final"
            if nota_final.value is not None:
                nota_final.value = round((media + nota_final.value) / 2, 1)
            else:
                nota_final.value = media  # Se nota_final for None, atribuir a média diretamente
        else:
            situacao.value = "Aprovado"
            nota_final.value = 0

    # Imprimir informações sobre o aluno
    print(f"Faltas: {faltas}, P1: {p1}, P2: {p2}, P3: {p3}, Situação: {situacao.value}, Exame Final: {nota_final.value}")

# Salvar as alterações na planilha
planilha.save("CópiadeEngenhariadeSoftware-DesafioCauêTamiaranav2.xlsx")
