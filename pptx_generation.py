import json
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Optional

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor


EMU_PER_INCH = 914400


# -----------------------------
# Data model
# -----------------------------
@dataclass
class Node:
    id: str
    name: str
    department: str
    title: str
    reports_to: Optional[str]
    children: List["Node"] = field(default_factory=list)
    depth: int = 0
    x_center: float = 0.0
    y_top: float = 0.0


# -----------------------------
# Build hierarchy
# -----------------------------
def build_tree(nodes_json: List[dict]) -> Node:
    nodes = {}
    for n in nodes_json:
        nodes[n["id"]] = Node(
            id=n["id"],
            name=n["name"],
            department=n["department"],
            title=n["title"],
            reports_to=n["reports_to"]
        )

    root = None
    for n in nodes.values():
        if n.reports_to and n.reports_to in nodes:
            nodes[n.reports_to].children.append(n)
        else:
            root = n

    if not root:
        raise ValueError("No root node found.")

    return root


def assign_depths(node: Node, depth: int = 0):
    node.depth = depth
    for c in node.children:
        assign_depths(c, depth + 1)


def collect_by_level(root: Node) -> Dict[int, List[Node]]:
    levels = defaultdict(list)

    def dfs(n: Node):
        levels[n.depth].append(n)
        for c in n.children:
            dfs(c)

    dfs(root)
    return levels


# -----------------------------
# Rendering helpers
# -----------------------------
def add_row_box(slide, x, y, w, h, text, bg_color, font_size, bold):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()

    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = 1  # center

    return shape


def draw_person(slide, node: Node, box_w, row_h):
    x = node.x_center - box_w / 2
    y = node.y_top

    colors = {
        "name": RGBColor(220, 230, 241),
        "dept": RGBColor(235, 241, 222),
        "title": RGBColor(242, 242, 242),
    }

    add_row_box(slide, x, y, box_w, row_h, node.name, colors["name"], 14, True)
    add_row_box(slide, x, y + row_h, box_w, row_h, node.department, colors["dept"], 12, False)
    add_row_box(slide, x, y + row_h * 2, box_w, row_h, node.title, colors["title"], 11, False)


def draw_connector(slide, parent: Node, child: Node, box_w, box_h):
    x1 = parent.x_center
    y1 = parent.y_top + box_h
    x2 = child.x_center
    y2 = child.y_top

    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1),
        Inches(y1),
        Inches(x2),
        Inches(y2),
    )
    conn.line.color.rgb = RGBColor(120, 120, 120)
    conn.line.width = Pt(1)


# -----------------------------
# Main rendering logic
# -----------------------------
def render_org_chart(
    root: Node,
    level_alignment: Dict[str, str],
    output_pptx: str
):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide_w = prs.slide_width / EMU_PER_INCH

    box_w = 2.6
    row_h = 0.4
    box_h = row_h * 3
    level_gap = 1.4
    margin = 0.6

    assign_depths(root)
    levels = collect_by_level(root)

    for depth, nodes in levels.items():
        count = len(nodes)
        total_width = count * box_w
        align = level_alignment.get(str(depth), "center")

        if align == "left":
            start_x = margin + box_w / 2
        elif align == "right":
            start_x = slide_w - margin - total_width + box_w / 2
        else:  # center
            start_x = (slide_w - total_width) / 2 + box_w / 2

        for i, n in enumerate(nodes):
            n.x_center = start_x + i * box_w
            n.y_top = margin + depth * level_gap

    for n in levels.values():
        for node in n:
            draw_person(slide, node, box_w, row_h)

    for depth_nodes in levels.values():
        for node in depth_nodes:
            for c in node.children:
                draw_connector(slide, node, c, box_w, box_h)

    prs.save(output_pptx)


# -----------------------------
# Entry point
# -----------------------------
if __name__ == "__main__":
    with open("pptx_json_flat.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    root = build_tree(data["nodes"])
    render_org_chart(
        root=root,
        level_alignment=data.get("level_alignment", {}),
        output_pptx="org_chart.pptx"
    )
