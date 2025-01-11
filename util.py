def decode_pos(pos):
    try:
        row = int(pos[1:]) - 1
        col = ord(pos[0]) - ord('A')
    except ValueError:
        row = -1
        col = -1

    return row, col


def encode_pos(row, col):
    return chr(col + ord('A')) + str(row + 1)
