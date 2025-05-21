def pad_row_to_length(row, length):
    return row[:length] + [''] * (length - len(row))
