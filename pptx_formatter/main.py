import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from pptx import Presentation
from langgraph.graph import END, StateGraph
from langgraph.pregel import Pregel


@dataclass
class SlideResult:
    index: int
    success: bool
    feedback: Optional[str] = None


def ai_judgement(state: dict) -> dict:
    """Placeholder AI judgement step."""
    # In real implementation, call an LLM or other model.
    state["ai_ok"] = True
    return state


def apply_template(state: dict) -> dict:
    """Copy the slide to a new presentation using the template."""
    template: Presentation = state["template"]
    slide = state["slide"]
    output: Presentation = state["output"]

    layout = template.slide_layouts[0]
    new_slide = output.slides.add_slide(layout)

    # Example: copy all shapes text from source slide
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        textbox = new_slide.shapes.add_textbox(left=shape.left, top=shape.top,
                                               width=shape.width, height=shape.height)
        textbox.text_frame.text = shape.text
    return state


def convert_to_jpeg(state: dict) -> dict:
    """Placeholder for JPEG conversion."""
    # Real implementation would export the slide to JPEG using python-pptx or win32com.
    return state


def quality_check(state: dict) -> dict:
    """Dummy quality check that always passes."""
    state["quality_ok"] = True
    return state


def mark_result(state: dict) -> SlideResult:
    feedback = None if state.get("quality_ok", False) else "Quality check failed"
    return SlideResult(index=state["index"], success=state.get("quality_ok", False), feedback=feedback)


def build_graph() -> Pregel:
    graph = StateGraph(dict)
    graph.add_node("ai", ai_judgement)
    graph.add_node("template", apply_template)
    graph.add_node("jpeg", convert_to_jpeg)
    graph.add_node("quality", quality_check)

    graph.set_entry_point("ai")
    graph.add_edge("ai", "template")
    graph.add_edge("template", "jpeg")
    graph.add_edge("jpeg", "quality")
    graph.add_edge("quality", END)

    return graph.compile()


def process_presentation(template_path: Path, target_path: Path, output_path: Path) -> List[SlideResult]:
    template = Presentation(template_path)
    target = Presentation(target_path)
    output = Presentation()

    results: List[SlideResult] = []
    graph = build_graph()

    for idx, slide in enumerate(target.slides, start=1):
        state = {"template": template, "slide": slide, "output": output, "index": idx}
        final_state = graph.invoke(state)
        results.append(mark_result(final_state))

    output.save(output_path)
    return results


def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(description="Generate a PPTX from template and target presentation")
    parser.add_argument("--template", required=True, type=Path, help="Path to template PPTX file")
    parser.add_argument("--target", required=True, type=Path, help="Path to target PPTX file")
    parser.add_argument("--output", type=Path, default=Path("output.pptx"), help="Destination PPTX file")
    args = parser.parse_args(argv)

    process_presentation(args.template, args.target, args.output)


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    main()
