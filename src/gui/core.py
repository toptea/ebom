from gooey import Gooey, GooeyParser

from gui.parser import cooperation
from gui.parser import electrical
from gui.parser import mechanical


options = {
    'program_name': 'BOM Import',
    'program_description': 'Version 0.4',
    'default_size': (610, 530),
    'navigation': 'SIDEBAR',
    'tabbed_groups': True,
    'image_dir': 'gui/img',
    'monospace_display': False,
}


@Gooey(**options)
def interface():
    """Gooey User Interface

    Turn any Python command line program into a full GUI application
    with one line. Follows argparse declaration format.

    Other Parameters
    ----------------
    program_name : str
        Defaults to script name
    program_description : str
        Defaults to ArgParse Description
    default_size : tuple
        starting size of the GUI
    navigation : str
        Sets the "navigation" style of Gooey's top level window.
        Choose between 'TABBED' or 'SIDEBAR'
    tabbed_groups : bool
        Tabbed the group instead of putting them in all in one window
    image_dir : str
        Path to the directory in which Gooey should look for
        custom images/icons
    monospace_display : bool
        Uses a mono-spaced font in the output screen

    Returns
    -------
    arg : obj
        return arguments
    """
    parser = GooeyParser()
    subs = parser.add_subparsers(help='commands', dest='command')

    subs = mechanical.add_parser(subs)
    # subs = electrical.add_parser(subs)
    # subs = cooperation.add_parser(subs)

    if subs is not None:
        args = parser.parse_args()
        args.open_model = not args.open_model
        args.recursive = not args.recursive
        args.update_drawing_revision = not args.update_drawing_revision
        args.update_parent_revision = not args.update_parent_revision
        args.exclude_same_revision = not args.exclude_same_revision
        args.output_report = not args.output_report
        args.output_import = not args.output_import
        return args
