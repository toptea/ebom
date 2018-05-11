import gui.core
import bom.core


def main():
    """Main entry point of the program

    Get the arguments using gui interface and then uses core.main()
    functions from their respective modules.
    """
    args = gui.core.interface()

    print('\nInput')
    for key, value in vars(args).items():
        print('|_', key, '=', value)

    bom.core.main(
        command=args.command,
        assembly=args.assembly,
        close_file=args.close_file,
        open_model=args.open_model,
        recursive=args.recursive,
        output_report=args.output_report,
        output_import=args.output_import
    )


if __name__ == '__main__':
    main()
