import argparse
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from threading import Lock
from typing import List, Optional

from PIL import Image, ImageDraw
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
    lock: Lock = state["lock"]

    layout = template.slide_layouts[0]
    with lock:
        new_slide = output.slides.add_slide(layout)

        # Example: copy all shapes text from source slide
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            textbox = new_slide.shapes.add_textbox(
                left=shape.left,
                top=shape.top,
                width=shape.width,
                height=shape.height,
            )
            textbox.text_frame.text = shape.text

    state["new_slide"] = new_slide
    return state


def convert_to_jpeg(state: dict) -> dict:
    """Render the new slide to a JPEG file."""
    slide = state["new_slide"]
    img_dir: Path = state["img_dir"]
    img_dir.mkdir(parents=True, exist_ok=True)
    img_path = img_dir / f"slide_{state['index']}.jpg"

    img = Image.new("RGB", (1280, 720), color="white")
    draw = ImageDraw.Draw(img)
    y = 10
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        draw.text((10, y), shape.text, fill="black")
        y += 20
    img.save(img_path, format="JPEG")

    state["img_path"] = img_path
    return state


def quality_check(state: dict) -> dict:
    """Simple quality check based on slide content."""
    slide = state["new_slide"]
    text_content = "\n".join(
        shape.text for shape in slide.shapes if shape.has_text_frame
    )
    state["quality_ok"] = "FAIL" not in text_content
    return state


def generate_feedback(state: dict) -> dict:
    state["feedback"] = f"Slide {state['index']} failed quality check"
    state["attempts"] = state.get("attempts", 0) + 1
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
    graph.add_node("feedback", generate_feedback)

    graph.set_entry_point("ai")
    graph.add_edge("ai", "template")
    graph.add_edge("template", "jpeg")
    graph.add_edge("jpeg", "quality")

    def route_quality(state: dict) -> str:
        return "ok" if state.get("quality_ok") else "fail"

    graph.add_conditional_edges(
        "quality", route_quality, {"ok": END, "fail": "feedback"}
    )

    def route_feedback(state: dict) -> str:
        return "retry" if state.get("attempts", 0) < 3 else "done"

    graph.add_conditional_edges(
        "feedback", route_feedback, {"retry": "ai", "done": END}
    )

    return graph.compile()


def process_single_slide(
    graph: Pregel,
    template: Presentation,
    output: Presentation,
    lock: Lock,
    img_dir: Path,
    idx: int,
    slide,
) -> SlideResult:
    state = {
        "template": template,
        "slide": slide,
        "output": output,
        "lock": lock,
        "img_dir": img_dir,
        "index": idx,
        "attempts": 0,
    }
    final_state = graph.invoke(state)
    return mark_result(final_state)


def process_presentation(template_path: Path, target_path: Path, output_path: Path) -> List[SlideResult]:
    template = Presentation(template_path)
    target = Presentation(target_path)
    output = Presentation()

    graph = build_graph()
    img_dir = output_path.parent / "images"
    lock = Lock()

    results: List[SlideResult] = []
    with ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(
                process_single_slide,
                graph,
                template,
                output,
                lock,
                img_dir,
                idx,
                slide,
            )
            for idx, slide in enumerate(target.slides, start=1)
        ]
        for fut in futures:
            results.append(fut.result())

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
