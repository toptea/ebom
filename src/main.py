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
    navigation='SIDEBAR'
)
def main():
    parser = GooeyParser()
    subs = parser.add_subparsers(help='commands', dest='command')
    subs = add_mech_parser(subs)
    subs = add_elec_parser(subs)
    subs = add_coop_parser(subs)

    arg = parser.parse_args()

    # Correct checkbox values
    arg.open_model = not arg.open_model
    arg.recursive = not arg.recursive

    print(arg)
    if arg.command == 'mechanical':
        bom.mechanical.save_csv_template(arg.partcode, arg.close_file, arg.open_model, arg.recursive)
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

    mech_parser.add_argument(
        '--partcode',
        type=str,
        help="enter the assembly you wish to import"
    )

    mech_parser.add_argument(
        '--open_model',
        action="store_false",
        help="if there's missing desc, open model"
    )

    mech_parser.add_argument(
        '--close_file',
        default='none',
        choices=['none', 'idw', 'iam', 'both'],
        help="close file when finished"
    )

    mech_parser.add_argument(
        '--recursive',
        action="store_false",
        help="generate sub assembly BOM"
    )
    return subs


def add_elec_parser(subs):
    elec_parser = subs.add_parser(
        'electrical',
        help='generate mechanical bom',
    )
    elec_parser.add_argument(
        'itemized_bom',
        widget="FileChooser",
        help="open promise report"
    )
    elec_parser.add_argument(
        '--raw',
        action="store_true",
        help="output raw data"
    )
    elec_parser.add_argument(
        '--current_bom',
        action="store_true",
        help="generate current bom only"
    )
    elec_parser.add_argument(
        '--close',
        action="store_true",
        help="close drawing when finished"
    )
    return subs


def add_coop_parser(subs):
    coop_parser = subs.add_parser(
        'cooperation',
        help='generate mechanical bom',
    )
    coop_parser.add_argument(
        'assembly',
        type=str,
        help="enter the partcode you wish to import"
    )
    coop_parser.add_argument(
        '--raw',
        action="store_true",
        help="output raw data"
    )
    coop_parser.add_argument(
        '--current_bom',
        action="store_true",
        help="generate current bom only"
    )
    coop_parser.add_argument(
        '--close',
        action="store_false",
        help="close drawing when finished"
    )
    return subs

if __name__ == '__main__':
    main()
