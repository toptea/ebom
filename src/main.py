import bom.mechanical
import bom.electrical
import bom.cooperation
import argparse
from gooey import Gooey, GooeyParser
import sys
import time


@Gooey(
    program_name='BOM Import',
    program_description='My Cool GUI Program!',
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

    # correct checkbox values
    if arg.command == 'mechanical':
        arg.open_model = not arg.open_model
        arg.recursive = not arg.recursive
        arg.output_raw_data = not arg.output_part_list

    print(arg)

    if arg.command == 'mechanical':
        bom.mechanical.main(arg.assembly, arg.close_file, arg.open_model, arg.recursive)
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
        'Options',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )
    save_group = mech_parser.add_argument_group(
        'Save',
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
        widget='Listbox',
        choices=['never', 'idw', 'iam', 'both'],
        help="Close inventor file when finished"
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


def add_elec_parser(subs):
    elec_parser = subs.add_parser(
        'electrical',
        help='generate mechanical bom',
    )

    bom_group = elec_parser.add_argument_group(
        'BOM',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )
    option_group = elec_parser.add_argument_group(
        'Options',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    save_group = elec_parser.add_argument_group(
        'Save',
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
        'BOM',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )
    option_group = coop_parser.add_argument_group(
        'Options',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    save_group = coop_parser.add_argument_group(
        'Save',
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
