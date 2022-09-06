import numpy as np
from openpyxl import load_workbook


def read_esheet(efile, esheet):  # считать лист из excel файла
    wb = load_workbook(efile)
    sheet = wb.get_sheet_by_name(esheet)
    return sheet


def read_range(esheet, range):  # считать диапозон ячеек с листа
    return esheet[range]


def range_to_list(erange):  # считать значения из ячеек
    elist =[]
    for erow in erange:
      for ecell in erow:
              elist.append(ecell.value)
    return elist


def range_to_dim(erange):  # создать список значений, считанных из ячеек
    edim = []
    elist =[]
    for erow in erange:
        for ecell in erow:
              elist.append(ecell.value)
        edim.append(elist)
        elist=[]
    return edim


def f(i):
    return int(i)


def dim_sotr_line(n_dim):  # сортировка по строкам
    a=[]
    c=[]
    for i in range (0, n_dim.shape[0]):
        a =( n_dim[i])
        c.append((sorted(a, key=f)))
    c=np.array(c)
    return(c)


def plus_list(a):  # создать список полученных из dim_sotr спиков 
    b = []
    for i in range (0, len(a)):
        f= a[i][0]
        b.append(f)
    return b


def dim_sotr(n_dim):  # сортировка по столбцам
    c=[]
    for i in range (0, n_dim.shape[1]):
        a =  n_dim[:,i:i+1]
        c.append((sorted(plus_list(a), key=f)))
    c=np.array(c)  
    c= np.transpose(c)
    return c


epath = './book.xlsx'
print("Введите имя листа (LIST1, LIST2 или LIST3)")   
eshit_name = input()
print("Введите диапозон значений (от A1 до E5 включительно) в формате A1:C2")
erange = input()
print("Введите направление сортировки")
print("1 - сортировка по столбцам") 
print("2 - сортировка по строкам")
sort_number = int(input())

esheet = read_esheet(epath, eshit_name)
elist = range_to_list(read_range(esheet,erange))
elist =sorted(elist,key= lambda x: x**3)
edim = range_to_dim(read_range(esheet,erange))
n_dim=np.array(edim)
print("Исходный массив"  ,n_dim,  sep='\n')

if sort_number==1:
    print("Введите параметр сортировки ( в порядке возрастания квадрата суммы корней 2-ой степени значений - ((int(x)**2)**0.5)*2)**2;  в порядке возрастания минимального значения-int(x) : ")
    fx = input()
    f = lambda x: eval(fx)
    print("Отсортированный массив", dim_sotr(n_dim), sep='\n')
elif sort_number==2:
    print("Отсортированный массив", dim_sotr_line(n_dim), sep='\n')
else:
    print("ОШИБКА")