import time
import os
import pandas as pd
import xlsxwriter
import natsort

# Trabalho de Projeto e Análise de Algoritmos
# Rafael Henrique da costa

# Padrão para cada instância:
# Primeira linha: número de objetos e capacidade da mochila
# Linhas seguintes: número do item, valor e peso de cada objeto
#
# Exemplo:
# N C
# 1 valor_1 peso_1
# 2 valor_2 peso_2
# N valor_N peso_N
#
# N = Número de objetos
# C = Capacidade da mochila


# Função de programação dinâmica para o problema da mochila
# values: lista de valores dos objetos
# weights: lista de pesos dos objetos
# capacity: capacidade da mochila
# retorna: valor máximo que pode ser colocado na mochila e lista de itens selecionados
def knapsack_dynamic_programming(values, weights, capacity):
    n = len(values)
    dp = [[0 for _ in range(capacity + 1)] for _ in range(n + 1)]

    for i in range(1, n + 1):
        for w in range(1, capacity + 1):
            if weights[i - 1] <= w:
                dp[i][w] = max(dp[i - 1][w], dp[i - 1][w - weights[i - 1]] + values[i - 1])
            else:
                dp[i][w] = dp[i - 1][w]

    selected_items = []
    i, w = n, capacity
    while i > 0 and w > 0:
        if dp[i][w] != dp[i - 1][w]:
            selected_items.append(i - 1)
            w -= weights[i - 1]
        i -= 1

    return dp[n][capacity], selected_items[::-1]


# Função com heurística gulosa para o problema da mochila
# values: lista de valores dos objetos
# weights: lista de pesos dos objetos
# capacity: capacidade da mochila
# retorna: valor máximo que pode ser colocado na mochila e lista de itens selecionados
def knapsack_greedy(values, weights, capacity):
    n = len(values)
    value_per_weight = [(values[i] / weights[i], i) for i in range(n)]
    value_per_weight.sort(reverse=True)

    total_value = 0
    selected_items = []
    remaining_capacity = capacity

    for _, index in value_per_weight:
        if weights[index] <= remaining_capacity:
            total_value += values[index]
            selected_items.append(index)
            remaining_capacity -= weights[index]

    return total_value, selected_items


# Função para ler a instância do problema da mochila
# file_path2: caminho do arquivo da instância
# retorna: lista de valores, lista de pesos, número de objetos e capacidade da mochila
def read_instance(file_path2):
    with open(file_path2, 'r') as f:
        lines = f.readlines()
        objects_number = int(lines[0].split()[0])
        values = [int(line.split()[0]) for line in lines]
        weights = [int(line.split()[1]) for line in lines]
        capacity = int(lines[0].split()[1])
        return values, weights, objects_number, capacity


# Função para resolver o problema da mochila e executar a comparação entre os algoritmos de programação dinâmica e
# heurística gulosa
# path: caminho da pasta com as instâncias
# dataframe: dataframe dos resultados
def knapsack_solve_execution(path, dataframe):
    os.chdir(path)
    lst = natsort.natsorted(os.listdir())
    for file in lst:
        if file.endswith(".txt"):
            file_path = f"{path}/{file}"
            values, weights, objects_number, capacity = read_instance(file_path)
            dynamic_start_time = time.time()
            max_value_dynamic, selected_items_dynamic = knapsack_dynamic_programming(values, weights, capacity)
            dynamic_execution_time = time.time() - dynamic_start_time
            greedy_start_time = time.time()
            max_value_greedy, selected_items_greedy = knapsack_greedy(values, weights, capacity)
            greedy_execution_time = time.time() - greedy_start_time
            new_row = {"Instance": file,
                       "Dynamic Execution Time": dynamic_execution_time,
                       "Greedy Execution Time": greedy_execution_time,
                       "Algorithms Execution Time Difference": greedy_execution_time - dynamic_execution_time,
                       "Maximum Value (Dynamic)": max_value_dynamic,
                       "Maximum Value (Greedy)": max_value_greedy,
                       "Selected Items (Dynamic)": selected_items_dynamic,
                       "Selected Items (Greedy)": selected_items_greedy}
            print("Instance: ", file)
            print("Execution time (Dynamic): ", dynamic_execution_time)
            print("Execution time (Greedy): ", greedy_execution_time)
            dataframe = pd.concat([dataframe, pd.DataFrame([new_row])], ignore_index=True)
    return dataframe


# Abaixo toda a lógica da extensão Pandas para exportar o dataframe para o Excel
directory = os.getcwd() + "/instancias"
df = pd.DataFrame(
    {
        "Instance": [],
        "Dynamic Execution Time": [],
        "Greedy Execution Time": [],
        "Algorithms Execution Time Difference": [],
        "Maximum Value (Dynamic)": [],
        "Maximum Value (Greedy)": [],
        "Selected Items (Dynamic)": [],
        "Selected Items (Greedy)": [],
    }
)

df = knapsack_solve_execution(directory, df)

writer = pd.ExcelWriter('../DynamicAndGreedyCompare.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='DynamicAndGreedyCompare', startrow=1, header=False, index=False)

workbook = writer.book
worksheet = writer.sheets['DynamicAndGreedyCompare']

(max_row, max_col) = df.shape

column_settings = [{'header': column} for column in df.columns]

worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

worksheet.set_column(0, max_col - 1, 30)

writer.close()
