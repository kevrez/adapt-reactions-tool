import xlrd

# save_path = r'"/Users/kevin/Dropbox/Python/Projects/ADAPT Reactions Tool/rxns.xls"'
RXN_PAGE = '(10)'

FIRST_RXN_ROW = 6
RXN_TYPE_COL = 2
RXN_VAL_COL = 3


DEAD_LOAD_TEXT = 'DL'
SDL_LOAD_TEXT = 'SDL'
LL_LOAD_TEXT = 'LL'

if __name__ == '__main__':
    DL = []
    SDL = []
    LL = []

    save_path = input('\nInput Excel Save Path. Note: Can include ":\n').replace('"', "")
    excel_report = xlrd.open_workbook(save_path)
    ws = excel_report.sheet_by_name(RXN_PAGE)

    seen_LL = False
    cur_row = FIRST_RXN_ROW

    try:
        while not seen_LL or ws.cell(cur_row, RXN_TYPE_COL).value:

            if ws.cell(cur_row, RXN_TYPE_COL).value == DEAD_LOAD_TEXT:
                DL.append(ws.cell(cur_row, RXN_VAL_COL).value)

            elif ws.cell(cur_row, RXN_TYPE_COL).value == SDL_LOAD_TEXT:
                SDL.append(ws.cell(cur_row, RXN_VAL_COL).value)

            elif ws.cell(cur_row, RXN_TYPE_COL).value == LL_LOAD_TEXT:
                LL.append(ws.cell(cur_row, RXN_VAL_COL).value)
                if not seen_LL:
                    seen_LL = True
            cur_row +=1

            if cur_row > 100:
                break

    except IndexError:
        print(f'\nReached end of file at row {cur_row}')

    if len(DL) == 0 or not (len(DL) == len(SDL) == len(LL)):
        print('There was an error with reading the reactions.')
    else:
        print("\nFormat:\nDL\nSDL\nLL\n")
        for i in range(len(DL)):
            print(f'Support {i+1}:')
            print(f'{DL[i]} k')
            print(f'{SDL[i]} k')
            print(f'{LL[i]} k')
            print()
    
    input('SUCCESS\nPress Enter to exit:\n')