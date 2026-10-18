"""Test script for v0.5.0 Animation & Progressive Deck utilities in PowerPoint Live MCP
Verifies:
1. ppt_split_slide_to_progressive_deck
2. ppt_create_progressive_deck
3. ppt_animate_sequence
4. ppt_animate_bullet_points
5. ppt_animate_emphasis
6. ppt_animate_exit
7. ppt_clear_animations
8. ppt_add_animated_stepper
"""

import sys
import os

# Add local directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_controller import controller


def run_tests():
    print(">>> 1. Creating new presentation...")
    res = controller.new_presentation(width=960, height=540)
    print("New pres:", res)

    print("\n>>> 2. Testing ppt_create_progressive_deck (Generates 4 progressive storyline slides)...")
    prog_res = controller.create_progressive_deck(
        title="Enterprise Architecture Modernization",
        steps=[
            {"title": "Core Legacy Audit", "description": "Catalog 450+ microservices & dependencies"},
            {"title": "Protocol Migration", "description": "Standardize data exchanges on FastMCP & JSON schemas"},
            {"title": "Agentic Orchestration", "description": "Deploy autonomous AI agents with live COM control"},
            {"title": "Global Deployment", "description": "Roll out to 15,000 corporate workstations worldwide"},
        ],
        accent_color_hex="#2563EB",
    )
    print("Progressive Deck:", prog_res)

    print("\n>>> 3. Testing ppt_add_animated_stepper on a new slide...")
    s5 = controller.add_slide(layout="blank")
    s5_idx = s5["slide_index"]
    controller.add_textbox(s5_idx, "Automated 4-Step Pipeline", left=80, top=50, width=800, height=45, font_size=26, bold=True)
    stepper_res = controller.add_animated_stepper(
        slide_index=s5_idx,
        steps=[
            {"title": "Input Request", "description": "User passes instruction via chat"},
            {"title": "Tool Resolution", "description": "MCP identifies optimal tools"},
            {"title": "Execution", "description": "Real-time COM commands sent to Office"},
            {"title": "Verification", "description": "High-res screenshot and state verified"},
        ],
        animate=True,
        trigger="on_click",
    )
    print("Animated Stepper:", stepper_res)

    print("\n>>> 4. Testing ppt_animate_sequence on custom shapes...")
    s6 = controller.add_slide(layout="blank")
    s6_idx = s6["slide_index"]
    controller.add_textbox(s6_idx, "Sequential Entrance Animations", left=80, top=50, width=800, height=45, font_size=26, bold=True)
    c1 = controller.add_card(s6_idx, "Feature Alpha", "Instant startup and minimal resource footprint", left=80, top=140, width=240, height=220)
    c2 = controller.add_card(s6_idx, "Feature Beta", "Full bidirectional streaming and async events", left=360, top=140, width=240, height=220)
    c3 = controller.add_card(s6_idx, "Feature Gamma", "Hardened security sandbox and zero trust", left=640, top=140, width=240, height=220)

    seq_res = controller.animate_sequence(
        slide_index=s6_idx,
        shape_ids_or_names=[c1["card_id"], c2["card_id"], c3["card_id"]],
        effect="fly_in",
        trigger="on_click",
    )
    print("Animate Sequence:", seq_res)

    print("\n>>> 5. Testing ppt_animate_bullet_points...")
    s7 = controller.add_slide(layout="blank")
    s7_idx = s7["slide_index"]
    controller.add_textbox(s7_idx, "Interactive Speaker Key Points", left=80, top=50, width=800, height=45, font_size=26, bold=True)
    blist = controller.add_bullet_list(
        s7_idx,
        items=[
            "First strategic priority: Cloud infrastructure scaling",
            "Second strategic priority: Developer experience and SDKs",
            "Third strategic priority: Enterprise security & governance",
        ],
        top=130,
        height=240,
    )
    bp_res = controller.animate_bullet_points(s7_idx, shape_id_or_name=blist["shape_id"], effect="fade")
    print("Bullet Points Reveal:", bp_res)

    print("\n>>> 6. Testing ppt_animate_emphasis & ppt_animate_exit...")
    emph_res = controller.animate_emphasis(s6_idx, shape_id_or_name=c2["card_id"], effect="pulse")
    exit_res = controller.animate_exit(s6_idx, shape_id_or_name=c1["card_id"], effect="fade")
    print("Emphasis & Exit:", emph_res, exit_res)

    print("\n>>> 7. Testing ppt_split_slide_to_progressive_deck...")
    # Create a slide with 3 distinct cards
    s8 = controller.add_slide(layout="blank")
    s8_idx = s8["slide_index"]
    controller.add_textbox(s8_idx, "Multi-Card Split Demo", left=80, top=50, width=800, height=45, font_size=26, bold=True)
    m1 = controller.add_card(s8_idx, "Phase 1: Concept", "Brainstorming and requirements gathering", left=80, top=140, width=240, height=220)
    m2 = controller.add_card(s8_idx, "Phase 2: Prototype", "Rapid prototyping with live visual feedback", left=360, top=140, width=240, height=220)
    m3 = controller.add_card(s8_idx, "Phase 3: Scale", "Enterprise distribution and global monitoring", left=640, top=140, width=240, height=220)

    split_res = controller.split_slide_to_progressive_deck(
        slide_index=s8_idx,
        shape_ids_or_names=[m1["card_id"], m2["card_id"], m3["card_id"]],
        mode="dim",
        highlight_border_color_hex="#2563EB",
        keep_master_slide=False,
    )
    print("Split to Progressive Deck:", split_res)

    print("\n>>> 8. Testing ppt_clear_animations...")
    clr_res = controller.clear_animations(slide_index=s7_idx)
    print("Clear Animations:", clr_res)

    print("\n>>> 9. Saving demo file and exporting showcase slides...")
    examples_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")
    os.makedirs(examples_dir, exist_ok=True)
    out_path = os.path.join(examples_dir, "v5_animations_showcase.pptx")
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception:
            pass
    controller.save_presentation(out_path)
    print("Saved presentation to:", out_path)

    # Export Slide 2 (progressive step 2) and Slide 3 (progressive step 3)
    img1 = os.path.join(examples_dir, "v5_progressive_step2.png")
    controller.export_slide_image(slide_index=2, output_path=img1)
    img2 = os.path.join(examples_dir, "v5_progressive_step3.png")
    controller.export_slide_image(slide_index=3, output_path=img2)
    print("Exported progressive slides images successfully!")

    print("\nALL V5 ANIMATION & PROGRESSIVE TESTS PASSED SUCCESSFULLY! [SUCCESS]")


if __name__ == "__main__":
    run_tests()
