# pptx-formatter

A minimal tool that uses `python-pptx` and `LangGraph` to build a new
PowerPoint presentation based on a template and an existing target
presentation. For each slide in the target file the tool copies all
textual content onto a new slide using the provided template. The
processing steps are represented as a LangGraph pipeline. Each slide is
analysed to determine its content type (text, image-heavy, or complex),
an appropriate template layout is selected from those available, and the
slide is rendered. Images of the original and formatted slides are
compared to detect problems; if an issue is found, feedback is generated
and a different template is tried, up to three attempts. Slides are
processed in parallel and JPEG previews are written to an `images`
folder alongside the output file.

## Installation

```bash
pip install .
```

## Usage

```bash
pptx-formatter --template template.pptx --target target.pptx --output out.pptx
```

This command produces `out.pptx` with slides taken from `target.pptx`
rendered in the style of `template.pptx`. JPEG previews for each slide
are placed in an `images` directory next to the output file. Slides
containing the word `FAIL` will trigger up to three retries before being
marked as failed.
