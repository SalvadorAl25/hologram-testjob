# 7. Escribe una función para filtrar números pares en una lista
import random

def filtrarPares(lista):
    return [num for num in lista if num%2==0]   #--- se retorna una lista de numeros si el numero de la lista es par

lista = []

for i in range(10):             #----> se itera con un rango de 10 elementos
    rand_num = random.randint(1,100)    # ---> se generan numeros aleatorios
    lista.append(rand_num)       #----> se agregan a la lista los numeros aleatorios

num_par = filtrarPares(lista)   # se ejecuta el metodo en base a la lista que recibe
print(f'lista de numeros aleatorios: {lista}')
print(f'los numeros pares de la lista de numeros aleatorios son: {num_par}')