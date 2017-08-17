from openpyxl import Workbook
import random

wb = Workbook()

# grab the active worksheet
ws = wb.active

#количество траифка по позициям
traffic = [1, 0.85, 0.75, 0.65, 0.06, 0.05, 0.04, 0.03, 0.02]

#задаем вероятности конверсий и количетсво трафика ключам
ver_conv = [0]*50
amount_traff = [0]*50
for i in range(49): 
    ver_conv[i] = round(random.uniform(0.1, 5)/100, 2)
    amount_traff[i] = random.randint(1, 100)
    
cost_click = [20]*50 #массив ставок по ключам. начальная ставка 20 руб.
base_cost_click = 20 #контрольная фиксированная ставка для всех ключей - 20 руб
cost_pos = [[0]*9 for i in range(50)] #массив стоимости позиций
all_count_conv = 0

for x in range (1, 10):
    all_count_conv = 0
    money = 0
    for v in range(1, 50):
        #генерируем стоимость позиций для ключа
        cost_pos[v][8] = round(random.uniform(1, 3), 2)
        cost_pos[v][7] = cost_pos[v][8] + round(random.uniform(0.5, 2), 2)
        cost_pos[v][6] = cost_pos[v][7] + round(random.uniform(0.5, 2), 2)
        cost_pos[v][5] = cost_pos[v][6] + round(random.uniform(0.5, 2), 2)
        cost_pos[v][4] = cost_pos[v][5] + round(random.uniform(0.5, 3), 2)
        cost_pos[v][3] = cost_pos[v][4] + round(random.uniform(4, 10), 2)
        cost_pos[v][2] = cost_pos[v][3] + round(random.uniform(2, 10), 2)
        cost_pos[v][1] = cost_pos[v][2] + round(random.uniform(1, 15), 2)
        cost_pos[v][0] = cost_pos[v][1] + round(random.uniform(2, 10), 2)

        #записываем в excel цены 
        for i in range(1, 10):
            ws.cell(row=v, column=i).value = cost_pos[v][i-1]

        #ищем максимальную доступную позицию и стоимость клика для оптимизатора
        cpc = [0]*50
        for i in range(9):
            if cost_pos[v][i] <= cost_click[v]:
                cpc[v] = cost_pos[v][i]+0.01
                num_pos = i
                break

        #максимальная позиция и стоимость клика для фиксированной ставки
        cpc_fix = 0
        for j in range(9):
            if cost_pos[v][j] <= base_cost_click:
                cpc_fix = cost_pos[v][j]+0.01
                num_pos_fix = j
                break
            
        #считаем количество трафика за период для оптимизатора
        traff_period = round(amount_traff[v] * traffic [num_pos], 0)

        #считаем количество трафика за период для фиксированной ставки
        traff_period_fix = round(amount_traff[v] * traffic [num_pos_fix], 0)

        #расходы по ключу оптимизатора
        costs_per_key = traff_period * cpc[v]

        #расходы по ключу фикс
        costs_per_key_fix = traff_period_fix * cpc_fix

        #количество конверсий
        count_conv = round(traff_period * ver_conv[v], 0)
        all_count_conv += count_conv

        conv_cost = [0]*50
        #стоимость конверсии по ключам
        if count_conv != 0:
            conv_cost[v] = costs_per_key / count_conv
            
        #расходы за период
        money += costs_per_key
        
        ws.cell(row=v, column=11).value = cpc[v] 
        ws.cell(row=v, column=12).value = num_pos
        ws.cell(row=v, column=13).value = traff_period
        ws.cell(row=v, column=14).value = costs_per_key
        ws.cell(row=v, column=15).value = count_conv
        ws.cell(row=v, column=16).value = conv_cost[v]

    #выставляем ставки
    for i in range(50):
        if (conv_cost[i] != 0) and (conv_cost[i] < money / all_count_conv):
            cost_click[i] += 2
            
            

ws.cell(row=50, column=15).value = all_count_conv

# Save the file
wb.save("sample1.xlsx")

print ("Done")
