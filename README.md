# pptx-formatter

A minimal tool that uses `python-pptx` and `LangGraph` to build a new
PowerPoint presentation based on a template and an existing target
presentation. For each slide in the target file the tool copies all
textual content onto a new slide using the provided template. The
processing steps are represented as a LangGraph pipeline, including
image generation, a simple quality check, and automatic retries. Each
slide is processed in parallel and exported as a JPEG to the `images`
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
