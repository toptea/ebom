from gooey import Gooey, GooeyParser

from gui.parser import cooperation
from gui.parser import electrical
from gui.parser import mechanical


options = {
    'program_name': 'BOM Import',
    'program_description': 'Version 0.3',
    'default_size': (610, 530),
    'navigation': 'SIDEBAR',
    'tabbed_groups': True,
    'image_dir': 'gui/img',
    'monospace_display': False,
}


@Gooey(**options)
def interface():
    parser = GooeyParser()
    subs = parser.add_subparsers(help='commands', dest='command')

    subs = mechanical.add_parser(subs)
    subs = electrical.add_parser(subs)
    subs = cooperation.add_parser(subs)

    if subs is not None:
        args = parser.parse_args()
        return args
