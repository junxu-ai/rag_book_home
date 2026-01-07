import json
import math
import argparse
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor


# -----------------------------
# Data model
# -----------------------------
@dataclass
class Node:
    id: str
    name: str
    title: str = ""
    reports_to: Optional[str] = None
    children: List["Node"] = field(default_factory=list)

    # layout fields
    depth: int = 0
    leaf_span: int = 1          # number of leaf nodes in subtree
    x_center: float = 0.0       # in EMU (we'll compute in inches then convert)
    y_top: float = 0.0


# -----------------------------
# Parsing: supports 2 JSON styles
# -----------------------------
def parse_json_to_tree(data: dict) -> Node:
    """
    Supports:
      A) Flat format:
         {"nodes":[{"id":"ceo","name":"A","title":"CEO","reports_to":null}, ...]}
      B) Nested format:
         {"id":"ceo","name":"A","title":"CEO","children":[{...}, {...}]}
         or {"root": {...}}
    Returns a single root Node.
    """
    if "root" in data and isinstance(data["root"], dict):
        return parse_nested_node(data["root"])

    if "nodes" in data and isinstance(data["nodes"], list):
        return parse_flat_nodes(data["nodes"])

    # If the entire JSON itself is a node dict (nested)
    if isinstance(data, dict) and "name" in data and ("children" in data or "id" in data):
        return parse_nested_node(data)

    raise ValueError(
        "Unsupported JSON format. Provide either {'nodes':[...]} flat format or a nested node with 'children'."
    )


def parse_nested_node(d: dict) -> Node:
    node_id = str(d.get("id") or d.get("name"))
    n = Node(
        id=node_id,
        name=str(d.get("name", "")),
        title=str(d.get("title", "")),
        reports_to=str(d.get("reports_to")) if d.get("reports_to") is not None else None,
    )
    for c in d.get("children", []) or []:
        child = parse_nested_node(c)
        child.reports_to = n.id
        n.children.append(child)
    return n


def parse_flat_nodes(nodes: list) -> Node:
    by_id: Dict[str, Node] = {}
    for p in nodes:
        pid = str(p["id"])
        by_id[pid] = Node(
            id=pid,
            name=str(p.get("name", "")),
            title=str(p.get("title", "")),
            reports_to=str(p["reports_to"]) if p.get("reports_to") is not None else None,
        )

    # build parent-child relationships
    roots: List[Node] = []
    for n in by_id.values():
        if n.reports_to and n.reports_to in by_id:
            by_id[n.reports_to].children.append(n)
        else:
            roots.append(n)

    if not roots:
        raise ValueError("No root found. Ensure at least one node has reports_to = null / missing.")
    if len(roots) > 1:
        # Create a synthetic root if multiple roots exist
        synthetic = Node(id="__root__", name="(Org Chart)", title="")
        for r in roots:
            r.reports_to = synthetic.id
            synthetic.children.append(r)
        return synthetic

    return roots[0]


# -----------------------------
# Layout algorithm (tree-based)
# -----------------------------
def assign_depths(root: Node, depth: int = 0) -> None:
    root.depth = depth
    for c in root.children:
        assign_depths(c, depth + 1)


def compute_leaf_spans(root: Node) -> int:
    """
    Leaf span = number of leaves in subtree.
    Used to allocate horizontal space proportionally and reduce overlap.
    """
    if not root.children:
        root.leaf_span = 1
        return 1
    s = 0
    for c in root.children:
        s += compute_leaf_spans(c)
    root.leaf_span = max(1, s)
    return root.leaf_span


def assign_positions(
    root: Node,
    x_left_in: float,
    x_right_in: float,
    y_top_in: float,
    level_gap_in: float,
) -> None:
    """
    Assign x_center (in inches) by distributing subtree spans across [x_left, x_right].
    """
    # y position determined purely by depth
    root.y_top = y_top_in + root.depth * level_gap_in

    # x center is midpoint of allocated band
    root.x_center = (x_left_in + x_right_in) / 2.0

    if not root.children:
        return

    total = sum(c.leaf_span for c in root.children)
    band_width = x_right_in - x_left_in
    cursor = x_left_in

    for c in root.children:
        w = band_width * (c.leaf_span / total if total else 1.0 / len(root.children))
        assign_positions(c, cursor, cursor + w, y_top_in, level_gap_in)
        cursor += w


def collect_nodes_edges(root: Node) -> Tuple[List[Node], List[Tuple[Node, Node]]]:
    nodes: List[Node] = []
    edges: List[Tuple[Node, Node]] = []

    def dfs(n: Node):
        nodes.append(n)
        for c in n.children:
            edges.append((n, c))
            dfs(c)

    dfs(root)
    return nodes, edges


# -----------------------------
# PPT rendering
# -----------------------------
def inches_to_emu(x_in: float) -> int:
    return int(Inches(x_in))


def add_person_box(
    slide,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    name: str,
    title: str,
    style: dict,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x_in),
        Inches(y_in),
        Inches(w_in),
        Inches(h_in),
    )

    # fill & line
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = style["fill_rgb"]

    line = shape.line
    line.color.rgb = style["line_rgb"]
    line.width = Pt(style["line_width_pt"])

    # text
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)

    # Name
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = name
    run.font.size = Pt(style["name_font_pt"])
    run.font.bold = True
    run.font.color.rgb = style["text_rgb"]

    # Title (optional)
    if title:
        p2 = tf.add_paragraph()
        run2 = p2.add_run()
        run2.text = title
        run2.font.size = Pt(style["title_font_pt"])
        run2.font.bold = False
        run2.font.color.rgb = style["text_rgb"]

    return shape


def add_report_line(slide, x1_in: float, y1_in: float, x2_in: float, y2_in: float, style: dict):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1_in),
        Inches(y1_in),
        Inches(x2_in),
        Inches(y2_in),
    )
    conn.line.color.rgb = style["connector_rgb"]
    conn.line.width = Pt(style["connector_width_pt"])
    return conn


def render_orgchart_to_pptx(
    root: Node,
    out_pptx: str,
    slide_w_in: float = 13.33,   # default widescreen
    slide_h_in: float = 7.5,
    margin_in: float = 0.5,
    box_w_in: float = 2.2,
    box_h_in: float = 0.85,
    level_gap_in: float = 1.2,
):
    prs = Presentation()
    prs.slide_width = Inches(slide_w_in)
    prs.slide_height = Inches(slide_h_in)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank slide

    style = {
        "fill_rgb": RGBColor(245, 247, 250),
        "line_rgb": RGBColor(120, 130, 140),
        "text_rgb": RGBColor(30, 35, 40),
        "line_width_pt": 1.0,
        "name_font_pt": 14,
        "title_font_pt": 11,
        "connector_rgb": RGBColor(120, 130, 140),
        "connector_width_pt": 1.0,
    }

    # layout
    assign_depths(root, 0)
    compute_leaf_spans(root)

    # starting y: reserve top margin
    y_top_in = margin_in
    x_left_in = margin_in
    x_right_in = slide_w_in - margin_in

    assign_positions(root, x_left_in, x_right_in, y_top_in, level_gap_in)
    nodes, edges = collect_nodes_edges(root)

    # map node id -> ppt shape
    shape_by_id: Dict[str, object] = {}

    # place boxes: convert center to top-left and keep within slide bounds
    for n in nodes:
        # skip synthetic root label if desired (here: we render it)
        x_in = n.x_center - box_w_in / 2.0
        y_in = n.y_top

        # clamp slightly to avoid falling off the slide
        x_in = max(margin_in, min(x_in, slide_w_in - margin_in - box_w_in))
        y_in = max(margin_in, min(y_in, slide_h_in - margin_in - box_h_in))

        shape = add_person_box(
            slide=slide,
            x_in=x_in,
            y_in=y_in,
            w_in=box_w_in,
            h_in=box_h_in,
            name=n.name,
            title=n.title,
            style=style,
        )
        shape_by_id[n.id] = shape

    # draw connectors
    # connector from manager bottom-center to report top-center
    for manager, report in edges:
        s1 = shape_by_id.get(manager.id)
        s2 = shape_by_id.get(report.id)
        if not s1 or not s2:
            continue

        x1 = (s1.left + s1.width / 2) / 914400.0  # EMU per inch
        y1 = (s1.top + s1.height) / 914400.0
        x2 = (s2.left + s2.width / 2) / 914400.0
        y2 = (s2.top) / 914400.0

        add_report_line(slide, x1, y1, x2, y2, style)

    prs.save(out_pptx)


# -----------------------------
# CLI
# -----------------------------
def main():
    ap = argparse.ArgumentParser(description="Generate a multi-level org chart PowerPoint from JSON.")
    ap.add_argument("--input", "-i", required=True, help="Path to input JSON file.")
    ap.add_argument("--output", "-o", required=True, help="Path to output PPTX file.")
    ap.add_argument("--landscape", action="store_true", help="Use widescreen landscape (default).")
    ap.add_argument("--portrait", action="store_true", help="Use portrait-like slide size.")
    args = ap.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    root = parse_json_to_tree(data)

    # slide sizing
    if args.portrait:
        slide_w_in, slide_h_in = 7.5, 13.33
    else:
        slide_w_in, slide_h_in = 13.33, 7.5

    render_orgchart_to_pptx(
        root=root,
        out_pptx=args.output,
        slide_w_in=slide_w_in,
        slide_h_in=slide_h_in,
    )


if __name__ == "__main__":
    main()
