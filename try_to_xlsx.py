from openpyxl import Workbook
import random

wb = Workbook()

# grab the active worksheet
ws = wb.active

#количество траифка по позициям
traffic = [1, 0.85, 0.75, 0.65, 0.06, 0.05, 0.04, 0.03, 0.02]

#задаем вероятности конверсий ключам
ver_conv = [0]*50
amount_traff = [0]*50
for i in range(49):
    ver_conv[i] = round(random.uniform(0.1, 5)/100, 2)
    amount_traff[i] = random.randint(1, 100)
    
cost_c = 15
cost_pos = [0]*9
all_count_conv = 0
for v in range(1, 50):
    #генерируем стоимость позиций для ключа
    cost_pos[8] = round(random.uniform(1, 3), 2)
    cost_pos[7] = cost_pos[8] + round(random.uniform(0.5, 2), 2)
    cost_pos[6] = cost_pos[7] + round(random.uniform(0.5, 2), 2)
    cost_pos[5] = cost_pos[6] + round(random.uniform(0.5, 2), 2)
    cost_pos[4] = cost_pos[5] + round(random.uniform(0.5, 3), 2)
    cost_pos[3] = cost_pos[4] + round(random.uniform(4, 10), 2)
    cost_pos[2] = cost_pos[3] + round(random.uniform(2, 10), 2)
    cost_pos[1] = cost_pos[2] + round(random.uniform(1, 15), 2)
    cost_pos[0] = cost_pos[1] + round(random.uniform(2, 10), 2)

    #записываем в excel цены 
    for i in range(1, 10):
        ws.cell(row=v, column=i).value = cost_pos[i-1]

    #ищем максимальную доступную позицию
    for i in range(9):
        if cost_pos[i] <= cost_c:
            cpc = cost_pos[i]+0.01
            num_pos = i
            break
        
    #считаем количество трафика за период
    traff_period = round(amount_traff[v] * traffic [num_pos], 0)

    #расходы по ключу
    costs_per_key = traff_period * cpc

    #количество конверсий
    count_conv = round(traff_period * ver_conv[v], 0)
    all_count_conv += count_conv

    conv_cost = 0
    #стоимость конверсии
    if count_conv != 0:
        conv_cost = costs_per_key / count_conv

    
    ws.cell(row=v, column=11).value = cpc 
    ws.cell(row=v, column=12).value = num_pos
    ws.cell(row=v, column=13).value = traff_period
    ws.cell(row=v, column=14).value = costs_per_key
    ws.cell(row=v, column=15).value = count_conv
    ws.cell(row=v, column=16).value = conv_cost

ws.cell(row=50, column=15).value = all_count_conv

    
# Save the file
wb.save("sample1.xlsx")
