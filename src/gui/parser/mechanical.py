from utils import system


def add_parser(subs):
    """Electrical BOM Parser

    Use to create UI for exporting electrical BOM.

    Returns
    -------
    subs : obj
        gooey/argparser sub-command
    """
    parser = subs.add_parser(
        'mechanical',
        help='generate mechanical bom',
    )

    parser = _add_bom_group(parser)
    parser = _add_option_group(parser)
    parser = _add_path_group(parser)
    parser = _add_save_group(parser)
    if parser is not None:
        return subs


def _add_bom_group(parser):
    """BOM Group"""
    group = parser.add_argument_group(
        'BOM',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )

    help_text_1 = "Enter the assembly number you wish to import.\n"
    help_text_2 = "If left blank, use the active drawing instead."
    group.add_argument(
        '--assembly',
        metavar='Assembly',
        type=str,
        help=help_text_1 + help_text_2
    )
    return parser


def _add_option_group(parser):
    """Option Group"""
    group = parser.add_argument_group(
        'Option',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    group.add_argument(
        '--open_model',
        metavar='Open Model',
        action="store_false",
        help="Fix missing description"
    )

    group.add_argument(
        '--close_file',
        metavar='Close',
        default='never',
        choices=['never', 'idw', 'iam', 'both'],
        help="Close file without saving"
    )

    group.add_argument(
        '--recursive',
        metavar='Recursive',
        action="store_false",
        help="Include sub assemblies"
    )

    group.add_argument(
        '--update_drawing_revision',
        metavar='Drawing Revision',
        action="store_false",
        help="Update drawing revision column"
    )

    group.add_argument(
        '--update_parent_revision',
        metavar='BOM Revision',
        action="store_false",
        help="Update parent revision column"
    )

    group.add_argument(
        '--exclude_same_revision',
        metavar='Exclude (WIP)',
        action="store_false",
        help="Exclude BOM with same revision"
    )
    return parser


def _add_path_group(parser):
    """Path Location Group"""
    group = parser.add_argument_group(
        'Location',
        gooey_options={
            'show_border': False,
            'columns': 1
        }
    )

    group.add_argument(
        '--vault_path',
        metavar='Meridian Vault (WIP)',
        default=str(system.find_vault_path()),
        widget="DirChooser",
        help="Where all the inventor files are stored"
    )

    group.add_argument(
        '--inventor_path',
        metavar='Autodesk Inventor (WIP)',
        default=str(system.find_inventor_path()),
        widget="FileChooser",
        help="The path location of the CAD program"
    )

    group.add_argument(
        '--output_path',
        metavar='Output (WIP)',
        default=str(system.find_export_path()),
        widget="DirChooser",
        help="Where the results are saved"
    )
    return parser


def _add_save_group(parser):
    """Save Group"""
    group = parser.add_argument_group(
        'Save',
        gooey_options={
            'show_border': False,
            'columns': 2
        }
    )

    group.add_argument(
        '--output_report',
        metavar='Report File',
        action="store_false",
        help="partcode.xlsx"
    )

    group.add_argument(
        '--output_import',
        metavar='Encompix Import',
        action="store_false",
        help="partcode.csv"
    )

    group.add_argument(
        '--output_item',
        metavar='New Item',
        action="store_true",
        help="partcode_item.csv"
    )
    return parser
