
number=["1ewdewfe"]
def to_number(value):
    
    try:
        return int(value)
    except (ValueError,TypeError):
        try:
            return float(value)
        except(ValueError,TypeError):
            return str(value)
print(to_number(number))
