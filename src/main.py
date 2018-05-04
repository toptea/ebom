import gui.core
import bom.core


def main():
    args = gui.core.interface()
    print(args)
    bom.core.main(
        command=args.command,
        assembly=args.assembly,
        close_file=args.close_file,
        open_model=args.open_model,
        recursive=args.recursive,
        output_report=True,
        output_import=True
    )


if __name__ == '__main__':
    main()
