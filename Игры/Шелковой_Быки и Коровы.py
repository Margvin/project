                    # Подключаем модели

from random import randint                    
import time
import os

                    # Создаем случайное число для игры

j = 0
while j != 1:
    X = randint(1000, 9999)
    x1 = X // 1000
    x2 = X % 1000 // 100
    x3 = X %100 // 10
    x4 = X % 10
    if x1 != x2 and x1 != x3 and x1 != x4:
        if x2 != x3 and x2 != x4:
            if x3 != x4:
                j = 1
                
                    # Текст для анимации
                    
s = '''             @@@@@@@@@                                                        @@@@@@@@@
           @@@@       @  @@                                              @@  @       @@@@
          @@           @@@@            @                    @            @@@@           @@
        @@            @@@ @@@          @@                  @@          @@@ @@@            @@
       @@            @@ @@    @         @@@              @@@         @    @@ @@            @@
       @              @  @@@@@@@@@@@     @@@            @@@     @@@@@@@@@@@  @              @
     @@@@@@           @@   @@@@   @@@@@ @@ @@          @@ @@ @@@@@   @@@@   @@           @@@@@@
    @@@@@@@@           @@@@ @@@@    @@@    @            @    @@@    @@@@ @@@@           @@@@@@@@
     @@@@@@@@        @  @@ @   @    @ @@@@@              @@@@@ @    @   @ @@  @        @@@@@@@@
     @@ @@@@@@    @@@@@@@  @@@@@   @@  @                    @  @@   @@@@@  @@@@@@@    @@ @@@@@@
     @@@@   @@@   @@@@@@@  @@@@@   @  @@@@@              @@@@@  @   @@@@@  @@@@@@@   @@@   @@@@
     @ @@    @@         @@@@@@@   @@@@@@@@@@            @@@@@@@@@@   @@@@@@@         @@    @@ @
     @@ @@    @@         @@@@@@@@@@@@@@ @@@@@          @@@@@ @@@@@@@@@@@@@@         @@    @@ @@
     @@  @@    @@@      @@@@@@    @@@@@     @@        @@     @@@@@    @@@@@@       @@@    @@  @@
     @@ @@@      @@@   @@ @        @@@@@                    @@@@@        @ @@   @@@      @@@ @@
     @@ @@        @    @@              @                    @              @@    @        @@ @@
     @@   @             @@@@          @@                    @@          @@@@             @   @@
     @@    @              @@@@@@ @@@@@@                      @@@@@@ @@@@@@              @    @@
     @@                    @@@@@@@@@                            @@@@@@@@@                    @@
     @                        @   @    @@@ @@@@@@  @@@@@@ @@@    @   @                        @
     @                        @@  @@      @@@@@@@  @@@@@@@      @@  @@                        @
     @                         @ @@  @@@@@@@@@@@@  @@@@@@@@@@@@  @@ @                         @
     @@                        @    @@    @@@@@      @@@@@    @@    @                        @@
      @@                        @@ @@      @            @      @@ @@                        @@
     @@@                         @@                              @@                         @@@    
    @                            @@                              @@                            @
    @                             @                              @                             @
  @@@@            @@              @             Быки             @              @@            @@@@
 @@@ @@            @@            @@                              @@            @@            @@ @@@
 @@@  @@@           @            @@              -               @@            @           @@@  @@@
 @@@@   @@@         @           @@                                @@           @         @@@   @@@@
@@  @@   @@@         @      @@@@ @             Коровы             @ @@@@      @         @@@   @@  @@
@    @   @@         @@    @@@@   @                                @   @@@@    @@         @@   @    @
@@  @@@   @@ @@    @@@@@@@    @@@                                  @@@    @@@@@@@    @@ @@   @@@  @@
 @  @@@    @@ @@@@@@ @@   @@@@                                        @@@@  @@   @@@@@@ @@   @@@  @
 @@ @       @  @@      @@    @@                                      @@    @@      @@   @      @ @@
  @@@      @@    @@      @@   @@@                                  @@@   @@      @@    @@      @@@
  @@       @@ @   @@@      @@@@@@@@@                            @@@@@@@@@      @@@   @ @@       @@
           @@@@@@@@@@@        @@@@@@@@                        @@@@@@@@        @@@@@@@@@@@
            @@@@@@@ @@          @@@@@@@                      @@@@@@@          @@ @@@@@@@
              @@@@@ @@                                                        @@ @@@@@'''

s1 = '''             @@@@@@@@@                                                        @@@@@@@@@
           @@@@       @  @@                                              @@  @       @@@@
          @@           @@@@            @                    @            @@@@           @@
        @@            @@@ @@@          @@                  @@          @@@ @@@            @@
       @@            @@ @@    @         @@@              @@@         @    @@ @@            @@
       @              @  @@@@@@@@@@@     @@@            @@@     @@@@@@@@@@@  @              @
      @               @@   @@@@   @@@@@ @@ @@          @@ @@ @@@@@   @@@@   @@               @
     @@                @@@@ @@@@    @@@    @            @    @@@    @@@@ @@@@                @@
     @               @  @@ @   @    @ @@@@@              @@@@@ @    @   @ @@  @               @
     @            @@@@@@@  @@@@@   @@  @                    @  @@   @@@@@  @@@@@@@            @
     @            @@@@@@@  @@@@@   @  @@@@@              @@@@@  @   @@@@@  @@@@@@@            @
     @                  @@@@@@@   @@@@@@@@@@            @@@@@@@@@@   @@@@@@@                  @
     @@                  @@@@@@@@@@@@@@ @@@@@          @@@@@ @@@@@@@@@@@@@@                  @@
     @@                 @@@@@@    @@@@@     @@        @@     @@@@@    @@@@@@                 @@
     @@                @@ @        @@@@@                    @@@@@        @ @@                @@
     @@                @@              @                    @              @@                @@
     @@   @             @@@@          @@                    @@          @@@@             @   @@
     @@   @@        @     @@@@@@ @@@@@@                      @@@@@@ @@@@@@     @        @@   @@
     @@   @@@      @@@     @@@@@@@@@                            @@@@@@@@@     @@@      @@@   @@
     @     @@@   @@@          @   @    @@@@ @@@@@@@@@@@@ @@@@    @   @          @@@   @@@     @
     @      @@  @@   @@@@     @@  @@       @@@@@@@@@@@@@@      @@  @@     @@@@   @@  @@      @
     @      @@@@ @@@     @@@   @ @@   @@@@@@@@@@@@@@@@@@@@@@@@  @@ @   @@@     @@@ @@@@      @
     @@            @@     @@@  @    @@     @@@@@    @@@@@    @@    @  @@@     @@            @@
      @@            @@ @@@@@@   @@ @@       @          @      @@ @@   @@@@@@ @@            @@
     @@@             @@@@@@@@    @@                              @@    @@@@@@@@             @@@    
    @                 @@@@@@@@   @@                              @@   @@@@@@@@@                @
    @                  @@@@@@     @                              @     @@@@@@                  @
  @@@@            @@              @             Быки             @              @@            @@@@
 @@@ @@            @@            @@                              @@            @@            @@ @@@
 @@@  @@@           @            @@              -               @@            @           @@@  @@@
 @@@@   @@@         @           @@                                @@           @         @@@   @@@@
@@  @@   @@@         @      @@@@ @             Коровы             @ @@@@      @         @@@   @@  @@
@    @   @@         @@    @@@@   @                                @   @@@@    @@         @@   @    @
@@  @@@   @@ @@    @@@@@@@    @@@                                  @@@    @@@@@@@    @@ @@   @@@  @@
@  @@@     @@ @@@@@@ @@  @@@@                                          @@@@  @@  @@@@@@ @@     @@  @
@@ @        @  @@     @@    @@                                        @@    @@     @@   @       @ @@
@@@        @@    @@    @@   @@@                                      @@@   @@    @@    @@        @@@
@@         @@ @   @@@  @@@@@@@@@                                    @@@@@@@@@  @@@   @ @@         @@
           @@@@@@@@@@@  @@@@@@@@                                    @@@@@@@@  @@@@@@@@@@@
            @@@@@@@ @@      @@@                                      @@@      @@ @@@@@@@
              @@@@@ @@                                                        @@ @@@@@'''

                    # Проигрышь начальной заставки
                    
n = 0
while n < 20:
    print(s)
    print(' ' * 45 + 'Загрузка - ' + str(n * 5) + '%')
    print(' '+ '/' * (5 * n) + '-' * (100 - 5 * n))
    n += 1
    time.sleep(0.01)
    os.system('CLS')
    print(s1)
    print(' ' * 45 + 'Загрузка - ' + str(n * 5) + '%')
    print(' '+ '/' * (5 * n) + '-' * (100 - 5 * n))
    n += 1
    time.sleep(0.01)
    os.system('CLS')
    
                    # Начальное меню - страница приветствия
                    
print(('*' * 100 + '\n' )* 2, end='')
print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
l = 'Довро пожаловать'
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = 'в игру'
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = '"Быки - Коровы" '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
print(('*' * 100 + '\n' )* 2, end='')
time.sleep(2)
os.system('CLS')
    
                    # Начальное меню - страница меню
                    
print(('*' * 100 + '\n' )* 2, end='')
print(('**' + ' ' * 96 + '**' + '\n') * 4, end='')
l = 'Если вы не знакомы с правилами данной игры, '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = 'или хотите освежить память, введите слово "Правила" '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = 'Если вы уже играли в нее, '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = 'введите слово "Старт", чтобы начать игру: '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
print('**', ' ' * 43, sep='', end='')
m = input()
print(('**' + ' ' * 96 + '**' + '\n') * 4, end='')
print(('*' * 100 + '\n' )* 2, end='')
time.sleep(1)
os.system('CLS')

                    # Провиряем введеную команду и отробатываем ее

j = 0
while j != 1:                  
    if m == "Правила" or m == 'правила':
    
                    # Начальное меню - страница "Правила"

        print(('*' * 100 + '\n') * 2, end='')
        print(('**' + ' ' * 96 + '**' + '\n') * 2, end='')
        l = '''Игра “Быки–Коровы” - логическая. Вы будете играть с компьютером, он загадывает
число из четырех неповторяющихся цифр (ноль в игре используется, но на
первом месте стоять не может), ваша же задача его угадать. Вы можете назвать любое
4-хзначное число, у которого цифры также не повторяются, оно сравниваться с загаданым.
При совпадении цифр в разрядах названного числа и загаданного, выводится сообщение “Бык”. 
Оно означает, что цифра отгадана и стоит в нужной позиции (например, в задуманном числе 
первая цифра 3 и в названном вами – тоже первая 3 – это бык.). "Корова" означает, 
что цифра отгадана, но она стоит не в своей позиции. Путем логических рассуждений и 
проверки ответов необходимо угадать все 4 цифры числа и их порядок. Ваша задача справииться 
за 10 попыток. '''
        n = len(l)
        c = 0
        while c < n:
            z = l.find('\n', c)
            e = len(l[c:z])
            probel = (int((96 - e) + 1) // 2)
            print('**', ' ' * probel, l[c:z], ' ' * probel, '**', sep='')
            c = c + e + 1
        print('**', ' ' * 26, sep='', end='')
        m = input('Введите "Старт" для запуска игры: ')
        print('*' * 100, '\n', '*' * 100, sep='')
        time.sleep(1)
        os.system('CLS')
    
                    # Запуск игры
                    
    elif m == 'Старт' or m == 'старт':
        j = 1
                    # Не верно введеный вариант меню
                    
    else:
        print(('*' * 100 + '\n' )* 2, end='')
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        l = 'Пункт меню введен не верно, '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        l = 'Пожалуйста, попробуйте еще раз: '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        print('**', ' ' * 43, sep='', end='')
        m = input()
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        print(('*' * 100 + '\n' )* 2, end='')
        time.sleep(1)
        os.system('CLS')

                    # Запрашиваем число у пользователя

print(('*' * 100 + '\n' )* 2, end='')
print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
l = 'Программа загадала целое число, из четырех неповторяющихся цифр '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
l = 'Пожалуйста, введите свой вариант числа: '
n = len(l)
probel = (int((96 - n) + 1) // 2)
print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
print('**', ' ' * 43, sep='', end='')
Y = input()
print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
print(('*' * 100 + '\n' )* 2, end='')
time.sleep(1)
os.system('CLS')

                        # Проверяем число на соответствие условиям

j = 0
while j != 1:
    if Y.isdigit() != True:

                        # Строка содержит символы, отличные от цифр
                        
        print(('*' * 100 + '\n' )* 2, end='')
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        l = 'Введеная вами строка содержит символы, отличные от цифр '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        l = 'Пожалуйста, введите свой вариант числа, из четырех неповторяющихся цифр:'
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        print('**', ' ' * 43, sep='', end='')
        Y = input()
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        print(('*' * 100 + '\n' )* 2, end='')
        time.sleep(1)
        os.system('CLS')

    elif len(str(Y)) != 4:
        
                    # Строка состоит не из 4 цифр
                        
        print(('*' * 100 + '\n' )* 2, end='')
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        l = 'Введеное вами число не 4-хзначное '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        l = 'Пожалуйста, введите целоe число, из четырех неповторяющихся цифр: '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        print('**', ' ' * 43, sep='', end='')
        Y = input()
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        print(('*' * 100 + '\n' )* 2, end='')
        time.sleep(1)
        os.system('CLS')
    
    #elif int(Y) // 1000 != int(Y) % 1000 // 100  and int(Y) // 1000 != int(Y) % 100 // 10 and int(Y) // 1000 != int(Y) % 10:
     #   if int(Y) % 1000 // 100 != int(Y) %100 // 10 and int(Y) % 1000 // 100 != int(Y) % 10:
     #       if int(Y) %100 // 10 != int(Y) % 10:
      #          j = 1
    else:

                    # Строка состоит повторяющихся цифр
                        
        print(('*' * 100 + '\n' )* 2, end='')
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        l = 'Введенок вами число состоит из повторяющихся цифр,'
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        l = 'Пожалуйста, введите целоe число, из четырех неповторяющихся цифр: '
        n = len(l)
        probel = (int((96 - n) + 1) // 2)
        print('**', ' ' * probel, l, ' ' * probel, '**', sep='')
        print('**', ' ' * 43, sep='', end='')
        Y = input()
        print(('**' + ' ' * 96 + '**' + '\n') * 5, end='')
        print(('*' * 100 + '\n' )* 2, end='')
        time.sleep(1)
        os.system('CLS')
        




                
print("asd asd asd as d")
input()
