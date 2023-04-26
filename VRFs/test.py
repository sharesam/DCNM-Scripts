a, b, c = 1, 2, 3


def f1():
    a, b, c = 11, 12, 13
    print(f'f1 scope for a is {a}, b is {b}, c is {c}')

    def f2():
        a, b, c = 21, 22, 23
        print(f'f2 scope for a is {a}, b is {b}, c is {c}')

    f2()
    print(f'updated f1 scope for a is {a}, b is {b}, c is {c}')


f1()
print('Global scope for a is {a}, b is {b}, c is {c}')
