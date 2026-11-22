"""Run the repository's Python generator regressions as one unittest suite.

Each file is loaded by path, so tests may sit beside the generator they cover
under any file name. Run through `pnpm test:python`, which supplies the
third-party dependency (fontTools) with uv.
"""

import importlib.util
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parent.parent
FILES = [
    ROOT / 'scripts/test_generate_font_reference_metrics.py',
]


def load(path: Path) -> unittest.TestSuite:
    name = path.name.removesuffix('.py').replace('-', '_').replace('.', '_')
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return unittest.defaultTestLoader.loadTestsFromModule(module)


if __name__ == '__main__':
    suite = unittest.TestSuite(load(path) for path in FILES)
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)
