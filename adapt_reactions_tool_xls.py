import xlrd

RXN_PAGE = '(10)'

FIRST_RXN_ROW = 6
RXN_TYPE_COL = 2
RXN_VAL_COL = 3

DEAD_LOAD_TEXT = 'DL'
SDL_LOAD_TEXT = 'SDL'
LL_LOAD_TEXT = 'LL'
MAX_SHEET_ROW = 50

if __name__ == '__main__':
    DL = []
    SDL = []
    LL = []

    while True:
        save_path = input('\nInput Excel Save Path. Note: Can include quotation marks ". Enter "exit" to end:\n').replace('"', '')
        if save_path.lower() in ('quit', 'exit', 'close', 'end'):
            break      
        
        seen_LL = False
        cur_row = FIRST_RXN_ROW

        try:
            excel_report = xlrd.open_workbook(save_path)
            ws = excel_report.sheet_by_name(RXN_PAGE)

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

                if cur_row > MAX_SHEET_ROW:
                    break

        except IndexError:
            print(f'\nReached end of file at row {cur_row}.')

        except FileNotFoundError:
            print('File not found.')
            continue

        if len(DL) == 0 or not (len(DL) == len(SDL) == len(LL)):
            print("Reactions not found. Ensure that the "
                + "'Moments, Shears, and Reactions' section is "
                + "included in the report.")
        else:
            print("\nFormat:\nDL\nSDL\nLL\n")
            for i in range(len(DL)):
                print(f'Support {i+1}:')
                print(f'DL = {DL[i]} k')
                print(f'SDL = {SDL[i]} k')
                print(f'LL = {LL[i]} k')
        