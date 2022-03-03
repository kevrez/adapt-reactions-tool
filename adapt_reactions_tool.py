from adapt_reactions_parser import AdaptReactionsParser

VERSION = "1.0"

WELCOME_MESSAGE = (f"ADAPT Reactions Tool V{VERSION} by Kevin Reznicek for U+C\n"
                   "This tool groups ADAPT reactions by supprt from .xls output.\n"
                   "Make sure your results come from runs with unskipped live loads.\n"
                   "Enter 'help' for instructions.\n")

HELP_MESSAGE = ("This program reads from ADAPT's .xls output.\n\n"
                "Once your standard design ADAPT run is complete, re-execute the "
                "run with the 'Skip Live Loads' option disabled.\n"
                "Set up an ADAPT output report including the 'Moments, Shears, "
                "and Reactions' tables under 'Tabular Reports - Compact'.\n"
                "Before creating the output file, make sure that the 'Create "
                "Optional XLS Report' option is checked.\n"
                "Save the report to a known location.\n\n"
                "Next, copy the file's full path by holding Shift "
                "and right-clicking the .xls file within Windows Explorer, "
                "then clicking 'Copy as Path'.\n"
                "Paste straight into the ADAPT Reactions Tool prompt.\n")


def get_path_from_input():
    print('Input Excel Save Path. Note: Can include "quotation marks".')
    print("Input 'q' to quit, 'h' for help:")
    save_path = input().replace('"', "")

    if save_path.lower() in ("q", "quit", "exit", "esc"):
        return "QUIT"
    if save_path.lower() in ("h", "help", "instructions", "howto"):
        return "HELP"
    return save_path


def main():
    while True:
        path = get_path_from_input().replace('""', '').strip()
        if path == "QUIT":
            break
        elif path.upper() == "HELP":
            print(HELP_MESSAGE)
        else:
            AdaptReactionsParser.print_reactions(path)


if __name__ == '__main__':
    print(WELCOME_MESSAGE)
    main()
