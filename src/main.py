import bom.mechanical
import bom.electrical
import bom.cooperation
from utils import system

from gooey import Gooey, GooeyParser
import argparse
import sys
import time


@Gooey(
    program_name='BOM Import',
    program_description='Version 0.3',
    default_size=(610, 530),
    navigation='SIDEBAR',
    tabbed_groups=True
)
def main():
    parser = GooeyParser()
    subs = parser.add_subparsers(help='commands', dest='command')
    subs = add_mech_parser(subs)
    subs = add_elec_parser(subs)
    subs = add_coop_parser(subs)

    arg = parser.parse_args()

    if arg.command == 'mechanical':
        arg.open_model = not arg.open_model
        arg.recursive = not arg.recursive
        arg.output_report = not arg.output_report
        arg.output_encompix = not arg.output_encompix

    if arg.command == 'mechanical':
        bom.mechanical.main(
            assembly=arg.assembly,
            close_file=arg.close_file,
            open_model=arg.open_model,
            recursive=arg.recursive,
            output_report=arg.output_report,
            output_encompix=arg.output_encompix
        )
    elif arg.command == 'electrical':
        bom.electrical.save_csv_templates()
    elif arg.command == 'cooperation':
        bom.cooperation.save_csv_template(arg.assembly)
    else:
        print('Invalid command')


def add_mech_parser(subs):
    mech_parser = subs.add_parser(
        'mechanical',
        help='generate mechanical bom',
    )

    bom_group = mech_parser.add_argument_group(
        'BOM',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )

    option_group = mech_parser.add_argument_group(
        'Option',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    path_group = mech_parser.add_argument_group(
        'Location',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )

    save_group = mech_parser.add_argument_group(
        'Save',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    help_text_1 = "Enter the assembly number you wish to import.\n"
    help_text_2 = "If left blank, use the active drawing instead."
    bom_group.add_argument(
        '-a',
        '--assembly',
        metavar='Assembly',
        type=str,
        help=help_text_1 + help_text_2
    )

    option_group.add_argument(
        '-o',
        '--open_model',
        metavar='Open Model',
        action="store_false",
        help="Fix missing description"
    )

    option_group.add_argument(
        '-c',
        '--close_file',
        metavar='Close',
        default='never',
        choices=['never', 'idw', 'iam', 'both'],
        help="Close file without saving"
    )

    option_group.add_argument(
        '-r',
        '--recursive',
        metavar='Recursive',
        action="store_false",
        help="Include sub assemblies"
    )

    option_group.add_argument(
        '-u',
        '--update_rev',
        metavar='Update (WIP)',
        action="store_false",
        help="Update BOM revision"
    )

    option_group.add_argument(
        '-x',
        '--exclude_same_rev',
        metavar='Exclude (WIP)',
        action="store_false",
        help="Exclude BOM with same revision"
    )

    path_group.add_argument(
        '-V',
        '--vault_path',
        metavar='Meridian Vault (WIP)',
        default=str(system.find_vault_path()),
        widget="DirChooser",
        help="Where all the inventor files are stored"
    )

    path_group.add_argument(
        '-I',
        '--inventor_path',
        metavar='Autodesk Inventor (WIP)',
        default=str(system.find_inventor_path()),
        widget="FileChooser",
        help="The path location of the CAD program"
    )

    path_group.add_argument(
        '-f',
        '--output_path',
        metavar='Output (WIP)',
        default=str(system.find_export_path()),
        widget="DirChooser",
        help="Where the results are saved"
    )

    save_group.add_argument(
        '-p',
        '--output_report',
        metavar='Summary Report',
        action="store_false",
        help="partcode.xlsx"
    )

    save_group.add_argument(
        '-e',
        '--output_encompix',
        metavar='Encompix',
        action="store_false",
        help="partcode.csv"
    )

    save_group.add_argument(
        '-i',
        '--output_item',
        metavar='Import1 (WIP)',
        action="store_true",
        help="partcode_item.csv"
    )

    return subs


def add_elec_parser(subs):
    elec_parser = subs.add_parser(
        'electrical',
        help='generate mechanical bom',
    )

    bom_group = elec_parser.add_argument_group(
        'BOM (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )
    option_group = elec_parser.add_argument_group(
        'Option (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    save_group = elec_parser.add_argument_group(
        'Save (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    bom_group.add_argument(
        'assembly',
        metavar='Assembly',
        widget="FileChooser",
        help="Choose the promise report you wish to import"
    )

    option_group.add_argument(
        '-x',
        '--exclude_same_rev',
        metavar='Exclude',
        action="store_false",
        help="Exclude BOM with same revision"
    )

    save_group.add_argument(
        '-p',
        '--output_part_list',
        metavar='Part List',
        action="store_true",
        help="partcode_raw.csv"
    )

    save_group.add_argument(
        '-e',
        '--output_encompix_csv',
        metavar='Encompix',
        action="store_false",
        help="partcode.csv"
    )

    save_group.add_argument(
        '-n',
        '--output_change_notice',
        metavar='Change Notice',
        action="store_false",
        help="partcode_cn.csv"
    )

    save_group.add_argument(
        '-l',
        '--output_log_file',
        metavar='Log File',
        action="store_false",
        help="partcode_log.csv"
    )
    return subs


def add_coop_parser(subs):
    coop_parser = subs.add_parser(
        'cooperation',
        help='generate mechanical bom',
    )

    bom_group = coop_parser.add_argument_group(
        'BOM (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )
    option_group = coop_parser.add_argument_group(
        'Option (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    save_group = coop_parser.add_argument_group(
        'Save (WIP)',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    bom_group.add_argument(
        '-a',
        '--assembly',
        metavar='Assembly',
        type=str,
        help="Enter the assembly you wish to import"
    )

    option_group.add_argument(
        '-r',
        '--recursive',
        metavar='Recursive',
        action="store_false",
        help="Include sub assemblies"
    )

    option_group.add_argument(
        '-x',
        '--exclude_same_rev',
        metavar='Exclude',
        action="store_false",
        help="Exclude BOM with same revision"
    )

    save_group.add_argument(
        '-p',
        '--output_part_list',
        metavar='Part List',
        action="store_true",
        help="partcode_raw.csv"
    )

    save_group.add_argument(
        '-e',
        '--output_encompix_csv',
        metavar='Encompix',
        action="store_false",
        help="partcode.csv"
    )

    save_group.add_argument(
        '-n',
        '--output_change_notice',
        metavar='Change Notice',
        action="store_false",
        help="partcode_cn.csv"
    )

    save_group.add_argument(
        '-l',
        '--output_log_file',
        metavar='Log File',
        action="store_false",
        help="partcode_log.csv"
    )
    return subs

if __name__ == '__main__':
    main()
