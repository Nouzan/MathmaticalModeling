

def fun(M, p, n):
    Mother = 1
    for i in range(int(M*p)+1):
        Mother *= (M-i)

    Child = 1
    for i in range(int(M*p)+1):
        Child *= (M-n-i)

    return Child/Mother


while True:
    M, p, n = input().split()
    print(fun(int(M), float(p), int(n)))
