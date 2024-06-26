import math
def count(n):
    squares = {i: i * i for i in range(1, n + 1)}
    count = 0

    for a in range(1, n):
        for b in range(a, n):
            c2 = squares[a] + squares[b]
            c = int(math.sqrt(c2))
            if c * c == c2 and c <= n:
                count += 1

    return count *2


# def count_mine(n):
#     c = 0
#     for s in range(1,n):
#         for b in range(s+1, n):
#             g = (s*s)+ (b*b)
#             l = str(math.sqrt(g)).split('.')
#             if int(l[1]) > 0:
#                 continue
#             c+=1
#     return c   

def count_mine(n):
    count = 0
    squares = {i: i * i for i in range(1, n + 1)}

    for a in range(1, n):
        for b in range(a, n):
            c2 = squares[a] + squares[b]
            c = int(math.sqrt(c2))
            if c * c == c2 and c <= n:
                count += 1

    return count *2


if __name__ == '__main__':
    print(count(5))
    print(count_mine(5))
    print(count(10))
    print(count_mine(10))
    print(f'Count {count(20)}')
    print(f'Count {count_mine(20)}')
    print(f'Count {count(40)}')
    print(f'Count {count_mine(40)}')
    print(f'Count {count(50)}')
    print(f'Count {count_mine(50)}')
    # print(count(30))
    # print(count(40))
    # print(count(50))
    
    