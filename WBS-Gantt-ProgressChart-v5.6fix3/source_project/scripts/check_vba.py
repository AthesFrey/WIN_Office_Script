"""Check VBA syntax using generated ANTLR VBA 7.1 grammar classes.

Generate the parser from grammars-v4/vba/vba7_1 with ANTLR 4.13.2, then pass
its directory via --parser-dir. This does not execute Excel or type-check COM.
"""
import argparse
import sys
from pathlib import Path

from antlr4 import CommonTokenStream, InputStream
from antlr4.error.ErrorListener import ErrorListener
from antlr4.atn.PredictionMode import PredictionMode
from antlr4.error.ErrorStrategy import BailErrorStrategy
from antlr4.error.Errors import ParseCancellationException


def check(source: Path, parser_dir: Path):
    sys.path.insert(0, str(parser_dir))
    from vbaLexer import vbaLexer
    from vbaParser import vbaParser

    class Errors(ErrorListener):
        def __init__(self):
            self.messages = []

        def syntaxError(self, recognizer, offendingSymbol, line, column, msg, e):
            self.messages.append(f"{source}:{line}:{column + 1}: {msg}")

    errors = Errors()
    # Exported document-module attributes are metadata, not VBA statements.
    # Keep blank lines so parser error line numbers still match the source.
    code = "\n".join("" if line.startswith("Attribute ") and not line.startswith("Attribute VB_Name =") else line
                     for line in source.read_text(encoding="utf-8").split("\n"))
    lexer = vbaLexer(InputStream(code))
    lexer.removeErrorListeners()
    lexer.addErrorListener(errors)
    tokens = CommonTokenStream(lexer)
    tokens.fill()
    parser = vbaParser(tokens)
    parser.removeErrorListeners()
    parser.buildParseTrees = False
    parser._interp.predictionMode = PredictionMode.SLL
    parser._errHandler = BailErrorStrategy()
    try:
        parser.startRule()
    except ParseCancellationException:
        # Full context is only needed for ambiguous/invalid VBA; the grammar's
        # default LL mode is prohibitively slow on long, valid procedure bodies.
        tokens.seek(0)
        parser = vbaParser(tokens)
        parser.buildParseTrees = False
        parser.removeErrorListeners()
        parser.addErrorListener(errors)
        parser.startRule()
    if errors.messages:
        raise SystemExit("\n".join(errors.messages))
    print(f"{source.name}: VBA 7.1 syntax passed (0 errors)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--parser-dir", type=Path, required=True)
    parser.add_argument("source", type=Path, nargs="?", default=Path(__file__).resolve().parents[1] / "src/ProgressChart.bas")
    args = parser.parse_args()
    check(args.source, args.parser_dir)
