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


def find_row_from_str(sheet, column, str_to_find):
    cur_row = 0
    while cur_row < 100:
        try:
            val = sheet.cell(cur_row, column).value
            if str(val).startswith(str_to_find):
                return cur_row
        except IndexError:
            print(f'Reached end of file and {str_to_find} was not found.')
            return None
        finally:
            cur_row += 1

    print('Looked at many rows and found nothing.')
    return None


def get_DL_SDL_rxns(sheet):
    DL = []
    SDL = []
    cur_row = find_row_from_str(sheet, 1, '5.2')
    try:
        while (len(SDL) == 0 or sheet.cell(cur_row, RXN_TYPE_COL).value) and (cur_row <= 100):
            if sheet.cell(cur_row, RXN_TYPE_COL).value == DEAD_LOAD_TEXT:
                DL.append(sheet.cell(cur_row, RXN_VAL_COL).value)
            elif sheet.cell(cur_row, RXN_TYPE_COL).value == SDL_LOAD_TEXT:
                SDL.append(sheet.cell(cur_row, RXN_VAL_COL).value)
            cur_row += 1
    except IndexError:
        print(f'\nReached end of file at row {cur_row}')
    return (DL, SDL)


def get_LL_rxns(sheet):
    LL = []
    cur_row = find_row_from_str(sheet, 1, '5.4') + 4
    is_unskipped = sheet.cell(cur_row, LL_RXN_VAL_COL).value == sheet.cell(
        cur_row, LL_RXN_VAL_COL+1).value
    try:
        while len(LL) == 0 or sheet.cell(cur_row, LL_RXN_VAL_COL).value:
            if sheet.cell(cur_row, LL_RXN_VAL_COL).value and is_unskipped:
                LL.append(sheet.cell(cur_row, LL_RXN_VAL_COL).value)
            cur_row += 1
    except IndexError:
        print(f'Reached EOF at row {cur_row}.')
    return LL


def print_reactions(DL: list[float], SDL: list[float], LL: list[float]):
    if not all((DL, SDL, LL)) or not len(DL) == len(SDL) == len(LL):
        # TODO: Move these checks to get_LL_rxns
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
            print(f'LL: {(LL[i] * RXN_FACTOR):.2f} k', '\n')


if __name__ == '__main__':

    save_path = input(
        '\nInput Excel Save Path. Note: Can include ":\n').replace('"', "")
    excel_report = xlrd.open_workbook(save_path)
    ws = excel_report.sheet_by_name(RXN_PAGE)

    dl_rxns, sdl_rxns = get_DL_SDL_rxns(ws)
    ll_rxns = get_LL_rxns(ws)

    print_reactions(DL=dl_rxns, SDL=sdl_rxns, LL=ll_rxns)
