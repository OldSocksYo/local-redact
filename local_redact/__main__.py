import sys

if len(sys.argv) > 1 and sys.argv[1] in {"--gui", "-g"}:
    from local_redact.gui import main

    raise SystemExit(main())

from local_redact.cli import main

raise SystemExit(main())
