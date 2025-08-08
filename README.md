# pptx-formatter

A minimal tool that uses `python-pptx` and `LangGraph` to build a new
PowerPoint presentation based on a template and an existing target
presentation. For each slide in the target file the tool copies all
textual content onto a new slide using the provided template. The
processing steps are represented as a LangGraph pipeline, allowing for
future extensions such as AI based quality checks.

## Installation

```bash
pip install .
```

## Usage

```bash
pptx-formatter --template template.pptx --target target.pptx --output out.pptx
```

This command produces `out.pptx` with slides taken from `target.pptx`
rendered in the style of `template.pptx`.
