import xlrd

RXN_PAGE = '(5)'

FIRST_RXN_ROW = 18
RXN_TYPE_COL = 2
RXN_VAL_COL = 3
LL_RXN_VAL_COL = 2

DL_RXN_FACTOR = 1
RXN_FACTOR = 1

DEAD_LOAD_TEXT = 'SW'
SDL_LOAD_TEXT = 'SDL'
LL_LOAD_TEXT = 'LL'


def find_row_from_str(ws, column, strToFind):
    cur_row = 0
    while cur_row < 100:
        try:
            val = ws.cell(cur_row, column).value
            # print(val)
            if str(val).startswith(strToFind):
                return cur_row
            cur_row += 1

        except IndexError:
            print(f'Reached end of file and {strToFind} was not found.')
            return None

    else:
        print('Looked at many rows and found nothing.')
        return None


if __name__ == '__main__':
    DL = []
    SDL = []
    LL = []

    save_path = input(
        '\nInput Excel Save Path. Note: Can include ":\n').replace('"', "")
    excel_report = xlrd.open_workbook(save_path)
    ws = excel_report.sheet_by_name(RXN_PAGE)

    cur_row = find_row_from_str(ws, 1, '5.2')

    # DL and SDL
    try:
        while (len(SDL) == 0 or ws.cell(cur_row, RXN_TYPE_COL).value) and (cur_row <= 100):
            if ws.cell(cur_row, RXN_TYPE_COL).value == DEAD_LOAD_TEXT:
                DL.append(ws.cell(cur_row, RXN_VAL_COL).value)
            elif ws.cell(cur_row, RXN_TYPE_COL).value == SDL_LOAD_TEXT:
                SDL.append(ws.cell(cur_row, RXN_VAL_COL).value)
            cur_row += 1
    except IndexError:
        print(f'\nReached end of file at row {cur_row}')

    # LL
    try:
        cur_row = find_row_from_str(ws, 1, '5.4') + 4
        while len(LL) == 0 or ws.cell(cur_row, LL_RXN_VAL_COL).value:
            isUnskipped = ws.cell(cur_row, LL_RXN_VAL_COL).value == ws.cell(
                cur_row, LL_RXN_VAL_COL+1).value
            if ws.cell(cur_row, LL_RXN_VAL_COL).value and isUnskipped:
                LL.append(ws.cell(cur_row, LL_RXN_VAL_COL).value)
            cur_row += 1
    except IndexError:
        print(f'Reached EOF at row {cur_row}.')

    # Output
    if len(DL) == 0 or not len(DL) == len(SDL) == len(LL):
        print('There was an error with reading the reactions.')
        print('Verify that the live loads in the file are not skipped.')
    else:
        print(
            f"\nFormat:\nDL\nSDL\nLL\n\nDL Reaction multiplied by {DL_RXN_FACTOR}")
        print(f"SDL and LL Reactions multiplied by {RXN_FACTOR}")
        for i, _ in enumerate(DL):
            print(f'Support {i+1}:')
            print(f'DL: {(DL[i] * DL_RXN_FACTOR):.2f} k')
            print(f'SD: {(SDL[i] * RXN_FACTOR):.2f} k')
            print(f'LL: {(LL[i] * RXN_FACTOR):.2f} k')
            print()
