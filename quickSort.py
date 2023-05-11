# Escribir un programa en python que implemente el algoritmo de ordenamiento quick sort

arr = [34, 7, 15, 9, 29, 1, 87, 45, 17, 10]

def quick_sort(arr):
    if len(arr) <= 1:   #----> si nada mas hay 1 elemento en el arreglo, retorna el valor
        return arr
    else:
        pivote = arr[0] #---> el primer valor del arreglo se guarda en otra variable que señalara el pivote
        menor = [n for n in arr[1:] if n <= pivote]   #---- se obtiene un arreglo con los elementos menores al pivote actual
        mayor = [n for n in arr[1:] if n > pivote]   # ----se obtiene un arreglo con los elementos mayores al pivote actual

    return quick_sort(menor) + [pivote] + quick_sort(mayor)  # ---> se aplica recursividad para realizar el ordenamiento de los demas elementos

lista_res = quick_sort(arr)
print(lista_res)

        