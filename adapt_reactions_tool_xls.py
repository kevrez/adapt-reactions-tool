import xlrd

RXN_PAGE = '(5)'

FIRST_RXN_ROW = 18
RXN_TYPE_COL = 2
RXN_VAL_COL = 3
LL_RXN_TYPE_COL = 1
LL_RXN_VAL_COL = 2

DEAD_LOAD_TEXT = 'SW'
SDL_LOAD_TEXT = 'SDL'
LL_LOAD_TEXT = 'LL'

DL_RXN_FACTOR = 1
RXN_FACTOR = 1


def find_row_from_str(sheet, column, str_to_find):
    cur_row = 0
    while cur_row < 200:
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


def get_worksheet_from_input():
    save_path = input(
        '\nInput Excel Save Path. Note: Can include ":\n').replace('"', "")
    excel_report = xlrd.open_workbook(save_path)
    worksheet = excel_report.sheet_by_name(RXN_PAGE)
    return worksheet


def get_rxn(sheet, rxn_type: str, row: int):
    rxn_col = {"DL": RXN_VAL_COL,
               "LL": LL_RXN_VAL_COL}
    selected_col = rxn_col[rxn_type]

    return sheet.cell(row, selected_col).value


def get_DL_SDL_rxns(sheet):
    DL = []
    SDL = []
    cur_row = find_row_from_str(sheet, column=1, str_to_find='5.2')

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
    ll_reactions = []
    cur_row = find_row_from_str(sheet, LL_RXN_TYPE_COL, '5.4') + 4

    ll_rxn_max = get_rxn(sheet, "LL", cur_row)
    ll_rxn_min = sheet.cell(cur_row, LL_RXN_VAL_COL).value
    is_run_skipped = ll_rxn_max != ll_rxn_min

    if is_run_skipped:
        raise Exception('Live Loads are not Skipped, verify your file input.')

    try:
        current_ll_rxn = None
        while len(ll_reactions) == 0 or current_ll_rxn:
            current_ll_rxn = get_rxn(sheet, "LL", cur_row)
            if current_ll_rxn:
                ll_reactions.append(current_ll_rxn)
            cur_row += 1
    except IndexError:
        print(f'Reached EOF at row {cur_row}.')

    return ll_reactions


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


def main():
    worksheet = get_worksheet_from_input()

    dl_rxns, sdl_rxns = get_DL_SDL_rxns(worksheet)
    ll_rxns = get_LL_rxns(worksheet)

    print_reactions(DL=dl_rxns, SDL=sdl_rxns, LL=ll_rxns)


if __name__ == '__main__':
    main()
