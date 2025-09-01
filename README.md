# PowerBB

PowerBB creates PowerPoint decks from a JSON description and includes an optional Qt based UI.

## Installation

Requires **Python 3.11**.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Command line

Run the builder directly:

```bash
python powerbb.py --help
```

Example to build a deck:

```bash
python powerbb.py --json input.json --output output.pptx
```

## GUI

Launch the graphical interface:

```bash
python powerbb_ui.py
```

## Slider CLI

Inspect existing PowerPoint files.

```bash
python slider.py --file samples/powerbb_roundtrip_test.pptx --list
python slider.py --file samples/powerbb_roundtrip_test.pptx --show 2
python slider.py --file samples/powerbb_roundtrip_test.pptx --json output/slides.json
```

## Testing

```bash
pytest -q
```

