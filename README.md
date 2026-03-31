# notebooklm-dewatermark

Remove the NotebookLM watermark from exported PPTX slides.

NotebookLM adds a small "NotebookLM" logo + text watermark to the bottom-right corner of every slide it generates. This tool removes it cleanly, handling any background color (white, black, blue, gradients, textures).

## How it works

NotebookLM exports each slide as a single full-page PNG image embedded in a PPTX file. The watermark is baked into the image itself. This tool:

1. Opens the PPTX and extracts each slide image
2. Samples the background pixels directly above the watermark area
3. Pastes the sampled strip over the watermark to seamlessly blend with the background
4. Saves a clean copy of the PPTX

## Install

```bash
git clone https://github.com/Longado/notebooklm-dewatermark.git
cd notebooklm-dewatermark
pip install -r requirements.txt
```

## Usage

```bash
# Single file
python3 notebooklm_dewatermark.py presentation.pptx

# Custom output name
python3 notebooklm_dewatermark.py presentation.pptx -o clean.pptx

# Batch processing
python3 notebooklm_dewatermark.py *.pptx
```

Output files are named `<original>_clean.pptx` by default.

## Before / After

| Before | After |
|--------|-------|
| ![before](assets/before.png) | ![after](assets/after.png) |

## Requirements

- Python 3.8+
- python-pptx
- Pillow

## License

MIT
