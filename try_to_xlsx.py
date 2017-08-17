from openpyxl import Workbook
import random

wb = Workbook()

# grab the active worksheet
ws = wb.active

#количество траифка по позициям
traffic = [1, 0.85, 0.75, 0.65, 0.06, 0.05, 0.04, 0.03, 0.02]

#ценность клика
click_value = 35

#задаем вероятности конверсий и количетсво трафика ключам
ver_conv = [0]*50
amount_traff = [0]*50
for i in range(50): 
    ver_conv[i] = round(random.uniform(0.1, 5)/100, 2)
    amount_traff[i] = random.randint(1, 100)
    
cost_click = [20]*50 #массив ставок по ключам. начальная ставка 20 руб.
base_cost_click = 20 #контрольная фиксированная ставка для всех ключей - 20 руб
cost_pos = [[0]*9 for i in range(50)] #массив стоимости позиций
all_count_conv = 0
all_money = 0
all_money_fix = 0
u = 0
l = 50

for x in range (1, 10):
    all_count_conv = 0
    all_count_conv_fix = 0
    money = 0
    money_fix = 0
    traff_period = [0]*50
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
            ws.cell(row=v+u, column=i).value = cost_pos[v][i-1]

        #ищем максимальную доступную позицию и стоимость клика для оптимизатора
        cpc = [0]*50
        for i in range(9):
            if cost_pos[v][i] <= cost_click[v]:
                cpc[v] = cost_pos[v][i]+0.01
                num_pos = i
                break

        #максимальная позиция и стоимость клика для фиксированной ставки
        cpc_fix = [0]*50
        for j in range(9):
            if cost_pos[v][j] <= base_cost_click:
                cpc_fix[v] = cost_pos[v][j]+0.01
                num_pos_fix = j
                break
            
        #считаем количество трафика за период для оптимизатора
        traff_period[v] = round(amount_traff[v] * traffic [num_pos], 0)

        #считаем количество трафика за период для фиксированной ставки
        traff_period_fix = round(amount_traff[v] * traffic [num_pos_fix], 0)

        #расходы по ключу оптимизатора
        costs_per_key = traff_period[v] * cpc[v]

        #расходы по ключу фикс
        costs_per_key_fix = traff_period_fix * cpc_fix[v]

        #количество конверсий оптимизатора
        count_conv = round(traff_period[v] * ver_conv[v], 0)
        all_count_conv += count_conv

        #количество конверсий с фикс ставкой
        count_conv_fix = round(traff_period_fix * ver_conv[v], 0)
        all_count_conv_fix += count_conv_fix

        conv_cost = [0]*50
        #стоимость конверсии по ключам с опитимизатором
        if count_conv > 0:
            conv_cost[v] = costs_per_key / count_conv

        conv_cost_fix = [0]*50
        #стоимость конверсии по ключам с фискированной ставкой
        if count_conv_fix > 0:
            conv_cost_fix[v] = costs_per_key_fix / count_conv_fix
            
        #расходы за период
        money += costs_per_key
        money_fix += costs_per_key_fix
        
        #фиксированная ставка 
        ws.cell(row=v+u, column=11).value = cpc_fix[v]
        ws.cell(row=v+u, column=12).value = num_pos_fix
        ws.cell(row=v+u, column=13).value = traff_period_fix
        ws.cell(row=v+u, column=14).value = costs_per_key_fix
        ws.cell(row=v+u, column=15).value = count_conv_fix
        ws.cell(row=v+u, column=16).value = conv_cost_fix[v]

        #оптимизированная ставка
        ws.cell(row=v+u, column=18).value = cpc[v]
        ws.cell(row=v+u, column=19).value = num_pos
        ws.cell(row=v+u, column=20).value = traff_period[v]
        ws.cell(row=v+u, column=21).value = costs_per_key
        ws.cell(row=v+u, column=22).value = count_conv
        ws.cell(row=v+u, column=23).value = conv_cost[v]


    #считаем расходы и конверсии у фиксированной ставки
    cost_conv_period_fix = money_fix / all_count_conv
    all_money_fix += money_fix

    #считаем расходы и конверсии у оптимизатора
    cost_conv_period = money / all_count_conv
    all_money += money

    ws.cell(row=l, column=14).value = money_fix
    ws.cell(row=l, column=21).value = money
    ws.cell(row=l, column=15).value = all_count_conv_fix
    ws.cell(row=l, column=22).value = all_count_conv
    ws.cell(row=l, column=16).value = cost_conv_period_fix
    ws.cell(row=l, column=23).value = cost_conv_period

    u += 52
    l += 52
    U = 0
    #выставляем ставки
    for i in range(50):
        U = traff_period[i] * ver_conv[i]

        if U * cpc[i] < click_value:
            cost_click[i] = cost_click[i] + cost_click[i]*0.1
        else:
            cost_click[i] = cost_click[i] - cost_click[i]*0.1
            
            

   

# Save the file
wb.save("sample1.xlsx")

print ("Done")
