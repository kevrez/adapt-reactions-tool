import xlrd


class AdaptReactionsParser:
    RXN_PAGE = '(5)'

    FIRST_RXN_ROW = 18
    RXN_TYPE_COL = 2
    RXN_VAL_COL = 3
    LL_RXN_TYPE_COL = 1
    LL_RXN_VAL_COL = 2

    DEAD_LOAD_TEXT = 'SW'
    SDL_LOAD_TEXT = 'SDL'
    LL_LOAD_TEXT = 'LL'

    @classmethod
    def _get_worksheet_from_path(cls, path: str):
        cleaned_path = path.replace('"', "")
        if not cleaned_path.lower().endswith(".xls"):
            raise FileNotFoundError("Received file was not a .xls file.")

        excel_report = xlrd.open_workbook(cleaned_path)
        # sheet_by_name will error with xlrd.biffh.XLRDError if there is no match
        worksheet = excel_report.sheet_by_name(cls.RXN_PAGE)
        return worksheet

    @classmethod
    def _find_row_from_str(cls, sheet, column: int, str_to_find: str):
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

    @classmethod
    def _get_rxn(cls, sheet, rxn_type: str, row: int):
        rxn_col = {"DL": cls.RXN_VAL_COL,
                   "LL": cls.LL_RXN_VAL_COL}
        selected_col = rxn_col[rxn_type]

        return sheet.cell(row, selected_col).value

    @classmethod
    def _get_DL_SDL_rxns(cls, sheet):
        DL = []
        SDL = []
        cur_row = cls._find_row_from_str(sheet, column=1, str_to_find='5.2')

        try:
            while (len(SDL) == 0 or sheet.cell(cur_row, cls.RXN_TYPE_COL).value) and (cur_row <= 200):
                if sheet.cell(cur_row, cls.RXN_TYPE_COL).value == cls.DEAD_LOAD_TEXT:
                    DL.append(sheet.cell(cur_row, cls.RXN_VAL_COL).value)
                elif sheet.cell(cur_row, cls.RXN_TYPE_COL).value == cls.SDL_LOAD_TEXT:
                    SDL.append(sheet.cell(cur_row, cls.RXN_VAL_COL).value)
                cur_row += 1
        except IndexError:
            if __name__ == '__main__':
                print(f'\nReached end of file at row {cur_row}\n')
        return (DL, SDL)

    @classmethod
    def _get_LL_rxns(cls, sheet):
        ll_reactions = []
        cur_row = cls._find_row_from_str(sheet, cls.LL_RXN_TYPE_COL, '5.4') + 4

        ll_rxn_max = cls._get_rxn(sheet, cls.LL_LOAD_TEXT, cur_row)
        ll_rxn_min = sheet.cell(cur_row, cls.LL_RXN_VAL_COL + 1).value
        is_run_skipped = ll_rxn_max != ll_rxn_min

        if is_run_skipped:
            raise ValueError(
                'Live Loads are not Skipped, verify your file input.')

        try:
            current_ll_rxn = None
            while len(ll_reactions) == 0 or current_ll_rxn:
                current_ll_rxn = cls._get_rxn(sheet, "LL", cur_row)
                if current_ll_rxn:
                    ll_reactions.append(current_ll_rxn)
                cur_row += 1
        except IndexError:
            if __name__ == '__main__':
                print(f'\nReached end of file at row {cur_row}.\n')

        return ll_reactions

    @classmethod
    def _reactions(cls, filename: str):
        sheet = cls._get_worksheet_from_path(filename)
        DL, SDL = cls._get_DL_SDL_rxns(sheet)
        LL = cls._get_LL_rxns(sheet)
        return DL, SDL, LL

    @classmethod
    def print_reactions(cls, filename: str):
        dl_rxns, sdl_rxns, ll_rxns = ([], [], [])
        try:
            dl_rxns, sdl_rxns, ll_rxns = cls._reactions(filename)
        except ValueError:
            print('\nLive Loads in the file are skipped and should not be.\n')
        except FileNotFoundError:
            print('\nReceived path did not point to .xls file.\n')
        except OSError:
            print('\nInvalid path! Ensure path starts with "P:" and file is .xls.\n')
        except xlrd.biffh.XLRDError:
            print("\nThe provided file doesn't contain reactions.\n"
                  "Make sure that 'Moments, Shears, and Reactions '"
                  "under 'Tabular Reports - Compact' is enabled.\n")

        if all((dl_rxns, sdl_rxns, ll_rxns)) and len(dl_rxns) == len(sdl_rxns) == len(ll_rxns):
            print()
            for i, _ in enumerate(dl_rxns):
                print(f'Support {i+1}:')
                print(f'DL: {(dl_rxns[i]):.2f} k')
                print(f'SD: {(sdl_rxns[i]):.2f} k')
                print(f'LL: {(ll_rxns[i]):.2f} k', '\n')


if __name__ == '__main__':
    SKIPPED_RUN_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Skipped_Live_Loads.xls"'
    UNSKIPPED_RUN_PATH = r'"C:\Users\kreznicek\Desktop\KR-Programs\Adapt-Reactions\test-reports\Valid_Reactions.xls"'

    print('*' * 50)
    print('SKIPPED RUN: SHOULD ERROR')
    AdaptReactionsParser.print_reactions(SKIPPED_RUN_PATH)

    print('*' * 50)
    print('UNSKIPPED RUN: SHOULD PRINT')
    AdaptReactionsParser.print_reactions(UNSKIPPED_RUN_PATH)
