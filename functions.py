

def find(lst, a):
    result = []
    for i, x in enumerate(lst):
        if x==a:
            result.append(i)
    return result