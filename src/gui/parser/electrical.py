

def add_parser(subs):
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