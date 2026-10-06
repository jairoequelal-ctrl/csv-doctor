import argparse
import json
from pathlib import Path
from .core import load_table, diagnose, apply, export


def main():
    parser = argparse.ArgumentParser(description='CSV Doctor — diagnose, preview and replay explicit repairs')
    sub = parser.add_subparsers(dest='command', required=True)
    ui = sub.add_parser('serve', help='Start local visual interface')
    ui.add_argument('--port', type=int, default=8765)
    diag = sub.add_parser('diagnose', help='Inspect a CSV/XLSX file without modifying it')
    diag.add_argument('file')
    diag.add_argument('--encoding', default='auto')
    diag.add_argument('--delimiter', default='auto')
    diag.add_argument('--sheet')
    clean = sub.add_parser('apply', help='Apply an explicit JSON recipe and export evidence')
    clean.add_argument('file')
    clean.add_argument('--recipe', required=True)
    clean.add_argument('--output', default='output')
    args = parser.parse_args()
    try:
        if args.command == 'serve':
            from .server import serve
            return serve(args.port)
        source = Path(args.file)
        if args.command == 'diagnose':
            table = load_table(source.read_bytes(), source.name, args.encoding,
                               '\t' if args.delimiter == 'tab' else args.delimiter, args.sheet)
            print(json.dumps(diagnose(table), indent=2, ensure_ascii=False))
        else:
            recipe = json.loads(Path(args.recipe).read_text(encoding='utf-8'))
            table = load_table(source.read_bytes(), source.name, **recipe.get('input', {}))
            result = apply(table, recipe)
            export(args.output, table, result)
            print(f"Exported {len(result['rows'])} rows; quarantined {len(result['quarantine'])}; "
                  f"{len(result['changes'])} changes; {len(result['unresolved'])} unresolved conversions.")
    except (ValueError, KeyError, TypeError, OSError) as exc:
        parser.exit(2, f'Error: {exc}\n')

if __name__ == '__main__':
    main()
