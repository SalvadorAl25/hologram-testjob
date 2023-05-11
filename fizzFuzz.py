# Escribir un programa en python que de solución al problema del FizzBuzz

for n in range(1,101):    #----> se Itera con el rango del 1 al 100
    if n %3 == 0 and n % 5 ==0:    #----> se verifica si es multiplo de 3 y de 5 
        print("FizzBuzz")
    elif n % 3 == 0:       #----> se verifica si es multiplo de 3
        print('Fizz')
    elif n % 5 == 0:        #----> se verifica si es multiplo de 5
        print('Buzz')
    else:                   # ---- si no, muestra el numero solamente
        print(n) 