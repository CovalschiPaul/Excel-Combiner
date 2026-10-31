# Project notes for automated agents

## What this project is

`excel_combiner_gui.py` is a single-file **Tkinter desktop GUI** (not a web app).
It reads every `.xlsx` file in a chosen folder, aligns the columns to a template
workbook, and writes one combined workbook. Dependencies: `pandas`, `openpyxl`
(see `requirements.txt`).

## Running it here

```bash
docker compose -f docker-compose.base44.yml up -d --build
```

Because the app is a desktop GUI, the stack runs a real X session in the
container and streams it to the browser:

- `Xvfb :99` — virtual display (1280x900), `openbox` as window manager
- `x11vnc` — serves that display on `127.0.0.1:5900`
- `websockify --web /usr/share/novnc` — port 6080, mapped to **host 3000**
- `preview/index.html` — a minimal noVNC viewer, copied over noVNC's
  `index.html` at startup so the preview root connects to the app directly

There is no hot reload for a Tk GUI: after editing the script run
`docker compose -f docker-compose.base44.yml restart excel-combiner`.

Source is mounted **read-only** at `/app`; `/workspace` is the writable working
directory, so `excel_combiner.log` and any combined output land there and never
in the repository. `sample_data/` holds a template plus two batch workbooks for
demoing the flow.

## Verified behaviour / gotchas

- Boot sequence: stale `/tmp/.X99-lock` makes Xvfb refuse to start, so the
  compose command clears it and waits until the display answers before starting
  the GUI. Without this, a container restart leaves a black screen.
- `x11vnc` must start only after the display is up, otherwise it exits.
- Driving the GUI headlessly (useful for checks): `xwininfo -id <win> -tree`
  gives the entry/button boxes; `xdotool mousemove X Y click 1` +
  `xdotool type "..."` fills the form. Entries are at x≈717, y≈362/414/466 and
  the Combine button at ≈640,540 on the default window placement.
- The combine run needs the Tk main thread free: the worker thread reads Tk
  variables and updates the status bar, so the thread pool must not be shut
  down inside the click handler (that deadlocks the app). Fixed by keeping a
  single long-lived executor created in `__init__`.
- Output of a demo run can be checked with
  `docker compose -f docker-compose.base44.yml exec -T excel-combiner /opt/venv/bin/python -c "import pandas as pd; print(pd.read_excel('/workspace/combined.xlsx'))"`.

## Preview-specific configuration

- The VNC session has no password (`-nopw`) because it is only reachable through
  the sandbox preview proxy; do not reuse this image for a public desktop.
- `BASE44_PREVIEW_MODE` is passed through to the service but no code branches on
  it — nothing in the app needs host/origin allowlisting (websockify serves
  static files and the websocket path is resolved from the page's own origin).
