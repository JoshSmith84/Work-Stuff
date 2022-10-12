

def remove_dup(a: list) -> None:
    """Remove duplicate strings from a given list.

    :param a: List to iterate through
    """
    i = 0
    while i < len(a):
        j = i + 1
        while j < len(a):
            if a[i] == a[j]:
                del a[j]
            else:
                j += 1
        i += 1

