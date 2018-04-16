import bom.mechanical
import bom.electrical
import bom.cooperation

def print_title():
    print(
        """
        --------------------------------------------------------------
          ____   ____  __  __   _____                            _
         |  _ \ / __ \|  \/  | |_   _|                          | |
         | |_) | |  | | \  / |   | |  _ __ ___  _ __   ___  _ __| |_
         |  _ <| |  | | |\/| |   | | | '_ ` _ \| '_ \ / _ \| '__| __|
         | |_) | |__| | |  | |  _| |_| | | | | | |_) | (_) | |  | |_
         |____/ \____/|_|  |_| |_____|_| |_| |_| .__/ \___/|_|   \__|
                                               | |
                                               |_|
        --------------------------------------------------------------
                                 By Gary Ip
        """
    )



def main_ui():
    print(
         """
        Welcome!

        This program provide the user an easy way to transfer data from
        Inventor or Promis-e, to the new ERP system Encompix. The data will
        be cleaned and saved on a csv file, which can be imported via
        the Multi-EBOM Import Utility.

        How to import the csv in to Encompix:
        a) Pick an option below and generate the csv file.
        a) In Encompix, Go to Inventory -> Custom -> Multi-EBOM Import Utility.
        c) Pick the csv you want to import and click Load/Validate.

        Please pick the following options:

        1) Generate Mechanical BOM
        2) Generate Electrical BOM
        3) Generate Cooperation BOM
        q) Quit
        """
    )

    while True:
        user_input = input('BOM Import: ')
        if user_input == '1':
            mech_bom_ui()
        elif user_input == '2':
            elec_bom_ui()
        elif user_input == '3':
            coop_bom_ui()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            print('Please enter a valid option')


def mech_bom_ui():
    print(
        """
        Generate Mechanical BOM csv
        ------------------------------------------------------------------

        This program will open an idw assembly drawing, extract the
        part list and save it to Encompix ebom template.
        If possible, it will also loop through all sub assemblies on
        the original part list. Please make sure to save and close
        all Inventor files before using this program.

        Please enter the assembly number you want to import:

        """
    )

    while True:
        user_input = input('Assembly: ')
        if user_input == 'b' or user_input == 'back':
            main_ui()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            bom.mechanical.save_csv_template(user_input)


def elec_bom_ui():
    print(
        """
        Generate Electrical BOM csv
        ------------------------------------------------------------------

        This program will loop through all "Bill of Materials - Itemised"
        reports in the current directory, and then convert them into
        Encompix csv template.

        Press 'Enter' to begin:

        """
    )

    while True:
        user_input = input('BOM Import: ')
        if user_input == 'b' or user_input == 'back':
            main_ui()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            bom.electrical.save_csv_templates()


def coop_bom_ui():
    print(
        """
        Generate Cooperation BOM csv
        ------------------------------------------------------------------

        This program will load the BOM from Cooperation Access database
        and then convert them into Encompix csv template.

        Please enter the assembly number you want to import:

        """
    )

    while True:
        user_input = input('Assembly: ')
        if user_input == 'b' or user_input == 'back':
            main_ui()
        elif user_input == 'q' or user_input == 'quit':
            exit()
        else:
            bom.cooperation.save_csv_template(user_input)


if __name__ == '__main__':
    print_title()
    main_ui()
