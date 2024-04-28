def is_number(str):
    try:
        float(str)
        return True
    except ValueError:
        return False

def is_empty(str):
    if str == '':
        return True
    else:
        return False