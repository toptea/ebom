

def add_parser(subs):
    """Cooperation BOM Parser

    Use to create UI for exporting Cooperation BOM.

    Returns
    -------
    subs : obj
        gooey/argparser sub-command
    """
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