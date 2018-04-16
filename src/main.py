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
    required_cols=1,
    optional_cols=2,
    navigation='SIDEBAR'
)
def main():
    parser = GooeyParser()

    subs = parser.add_subparsers(help='commands', dest='command')

    mech_parser = subs.add_parser(
        'mechanical',
        help='generate mechanical bom',
    )
    mech_parser.add_argument(
        'assembly',
        type=str,
        help="enter the partcode you wish to import"
    )
    mech_parser.add_argument(
        '--raw',
        action="store_true",
        help="output raw data"
    )
    mech_parser.add_argument(
        '--current_bom',
        action="store_true",
        help="generate current bom only"
    )
    mech_parser.add_argument(
        '--close',
        action="store_true",
        help="close drawing when finished"
    )

    elec_parser = subs.add_parser(
        'electrical',
        help='generate mechanical bom',
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

    arg = parser.parse_args()
    print(arg)
    if arg.command == 'mechanical':
        bom.mechanical.save_csv_template(arg.assembly)
    elif arg.command == 'electrical':
        bom.electrical.save_csv_templates()
    elif arg.command == 'cooperation':
        bom.cooperation.save_csv_template(arg.assembly)
    else:
        print('Invalid command')


if __name__ == '__main__':
    main()
