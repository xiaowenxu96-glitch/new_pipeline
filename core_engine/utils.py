def column_letter_to_number(column_letter):
    """将Excel列字母转换为数字索引（A->1, B->2, AA->27等）"""
    result = 0
    for char in column_letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result

def column_number_to_letter(column_number):
    """将数字索引转换为Excel列字母（1->A, 2->B, 27->AA等）"""
    result = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(ord('A') + remainder) + result
    return result
