import tkinter as tk  
from tkinter import *
import tkinter.ttk as ttk
import json
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import re
import pdb


def update_combo_wells(field):
	field = combo_field.get()
	well_list = [x for x in input_data[field].keys()]
	combo_well['values'] = well_list
	combo_well.delete(0, tk.END)
	
def labels(x, y):
	for x,y in zip(x, y):
		label = "{:.2f}".format(y)
		plt.annotate(label, #текст
					 (x,y), # координаты
					 textcoords="offset points", # расположение текста
					 xytext=(0,10), # расстояние между точками
					 ha='center') # выравнивание
 
# def getting_zakachka(text):
	# pdb.set_trace()
	# print(text)
	# try:
		# zakachka = re.search(r'чка\s*([^м]+)м', text['Примечание'])
		# print(zakachka)
		# number = zakachka.group(1).strip()
		# print(number)
		# text['Примечание'] = int(number)
		# return text['Примечание']
	# except Exception as e:
		# print(text)
		# text['Примечание'] = None
		# return text['Примечание']

def Plotting():
	plt.clf()
	day_on_fond = False
	legend_mark = ''
	
	well_df = pd.read_excel(r'C:\Users\Shakin.VYu\Desktop\База данных ВНР - чистовик.xlsx')
	well_df.columns = ['Дата', 'Скважина', 'Куст', 'Месторождение', 'Вид ремонта', 'Кол-во дней пред. ВНР', 'Остан. Дебит жидк.', 
	'Остан. обводн.', 'Остан. дебит нефти', 'Ожид. дебит жидк.', 'Ожид. обводн.', 'Ожид. дебит нефти', 
	'Прогн. дебит жидк.', 'Прогн. обводн.', 'Прогн. дебит нефти', 'Текущ. дебит жидк.', 'Текущ. обводн.', 
	'Текущ. дебит нефти', '+/- К предыд. суткам, жидкость', '+/- К предыд. суткам, нефть', 'Недостижение, жидк.', 
	'Недостижение, нефть', 'Дата запуска на ВНР', 'Дни на ВНР', 'Дата запуска по фонду, график', 'Прогнозная дата запуска по фонду', 
	'Давление на приеме', 'Отбор, м3', 'Примечание', 'Метка запуска по фонду']
	well_df['Примечание'] = well_df['Примечание'].str.extract('(\d+)').astype(float)
	well_df = well_df.loc[well_df['Текущ. обводн.']!= '-']
	well_df['Скважина'] =  well_df['Скважина'].astype('string')
	well_df['Текущ. обводн.'] =  well_df['Текущ. обводн.']*100
	
	well_number = combo_well.get()
	
	well_df = well_df[well_df['Скважина'] == well_number]
	well_df['Процент отбора'] = well_df['Отбор, м3']/well_df['Примечание']*100
	well_df['Процент отбора'].values[-1] = well_df['Процент отбора'].values[-2]
	if any(well_df['Метка запуска по фонду'].values == 'Запуск по фонду'):
		df_on_fond = well_df[well_df['Метка запуска по фонду'] == "Запуск по фонду"]
		day_on_fond = df_on_fond['Дата'].values[-1]
	else:
		mark_on_fond = '(По фонду не запущена)'
		
	plt.bar(well_df['Дата'], well_df['Процент отбора'], label ='Процент отбора', linewidth = 1)
	plt.plot(well_df['Дата'], well_df['Текущ. дебит жидк.'], color = 'orange', label ='Дебит жидкости, м3', marker ='o', linewidth = 1)
	plt.plot(well_df['Дата'], well_df['Текущ. обводн.'], color = 'brown', label ='Обводненность, %', marker ='o', linewidth = 1)
	labels(well_df['Дата'], well_df['Текущ. обводн.'])
	labels(well_df['Дата'], well_df['Процент отбора'])
	labels(well_df['Дата'], well_df['Текущ. дебит жидк.'])
	plt.xticks(rotation=45, ha="right")
	
	if any(well_df['Дата запуска по фонду, график'].values !='вне графика'):
		df2=well_df.loc[well_df['Дата запуска по фонду, график']!= 'вне графика']
		day_on_fond_graph = df2['Дата запуска по фонду, график'].values[-1].strftime('%d-%m-%y').replace('-','.')
		plt.axvline(x = df2['Дата запуска по фонду, график'].values[-1].strftime('%d-%m-%y').replace('-','.'), color = 'green', label = f'Дата запуска по фонду (график):\n{day_on_fond_graph}', linewidth = 1)
		legend_mark = ''
		
	else:
		legend_mark = 'Вне графика'
	if day_on_fond:
		plt.axvline(x = str(well_df['Дата'].values[-1]), color = 'red', label = f'Дата запуска по фонду (факт):\n{day_on_fond}', linewidth = 1)
	else:	
		legend_mark += mark_on_fond


	plt.xlabel('Дата', fontsize=16)
	plt.title(f'Изменение в период ВНР дебита по скважине {well_number} {legend_mark}')
	plt.grid(color = 'blue',linestyle ='--', linewidth = 0.5)
	plt.legend(loc='best')
	plt.rcParams['lines.linewidth'] = 8
	
	plt.show()

	
	
with open(r'C:\Users\Shakin.VYu\Desktop\code\Приложение ВНР\VNR.json', encoding='windows-1251') as file:
	input_data = json.load(file)   
input_data = input_data


app = tk.Tk()


screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
window_width = 500
window_height = 300
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

app.geometry(f"{window_width}x{window_height}+{x}+{y}")
app.title("Приложение ВНР")

combo_field = ttk.Combobox(values=list(input_data.keys()))
combo_field.bind("<<ComboboxSelected>>", update_combo_wells)
combo_field.grid(row=1, column=1, padx=1, pady=1)

combo_well = ttk.Combobox(values=[])
combo_well.grid(row=2, column=1, padx=1, pady=1)

btn_load = ttk.Button(text = 'Выгрузить график', command = Plotting)
btn_load.grid(row=3, column=1, padx=1, pady=1)

# Запускаем программу
app.mainloop()

 