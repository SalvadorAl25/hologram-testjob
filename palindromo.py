# Escribir un programa en python que permita saber si una palabra es palíndromo

pal = input('Inserte una palabra: ')
pal_inv = pal[::-1]  #---- se invierte el orden de los elementos y se asigna a un valor

if pal == pal_inv:    # --- se verifica si las cadenas son iguales
    print(f'{pal} es un palíndromo')
else:
    print(f'{pal} no es un palíndromo')