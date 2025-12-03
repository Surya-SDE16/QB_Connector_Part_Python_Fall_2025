from __future__ import annotations  # Postponed annotations for consistency

import sys  # Access to sys.exit

from .qb_cli import main  # Import CLI entrypoint

if __name__ == "__main__":  # pragma: no cover
    # Allow running as a module: python -m payment_terms_cli
    sys.exit(main())
