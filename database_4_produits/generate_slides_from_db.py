"""Generates a PowerPoint presentation live on screen from the 4 products database!
Demonstrates combining the database with the PowerPoint Live Controller.
"""

import os
import sys
import time

# Import database and PowerPoint controller
sys.path.append(os.path.join(os.path.dirname(__file__), "..", "powerpoint-mcp"))
from ppt_controller import controller
from init_db import get_all_products


def generate_presentation():
    products = get_all_products()
    print(f">>> Found {len(products)} products in SQLite database.")
    print(">>> Launching PowerPoint and creating catalog presentation live...")

    # 1. New presentation
    controller.new_presentation(width=960, height=540)
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 1: Catalog Overview (Grid of 4 Products)
    # ----------------------------------------------------
    print(">>> Building Slide 1: Catalog Grid...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(1, "#0F172A", gradient_color_hex="#1E293B")
    time.sleep(0.5)

    controller.add_textbox(
        slide_index=1,
        text="Catalogue Produits High-Tech 2026",
        left=80,
        top=40,
        width=800,
        height=50,
        font_size=30,
        bold=True,
        color_hex="#F8FAFC",
    )
    controller.add_textbox(
        slide_index=1,
        text="Généré automatiquement depuis la base de données SQLite (database_4_produits)",
        left=80,
        top=90,
        width=800,
        height=30,
        font_size=14,
        color_hex="#94A3B8",
    )
    time.sleep(0.5)

    # 4 Product Cards side-by-side (2x2 or 4 horizontal)
    # Let's do 4 horizontal sleek cards
    card_w = 185
    card_h = 320
    gap = 20
    start_x = 80
    start_y = 140

    accents = ["#3B82F6", "#10B981", "#8B5CF6", "#F59E0B"]

    for i, p in enumerate(products):
        cx = start_x + i * (card_w + gap)
        body_text = f"Prix: ${p['price']:.2f}\nStock: {p['stock_quantity']} pcs\nNote: {p['rating']} ★\n\n{p['short_description'][:75]}..."
        controller.add_card(
            slide_index=1,
            title=p["brand"] + "\n" + p["name"][:25] + "...",
            body=body_text,
            left=cx,
            top=start_y,
            width=card_w,
            height=card_h,
            bg_color_hex="#FFFFFF",
            border_color_hex="#334155",
            title_color_hex="#0F172A",
            body_color_hex="#475569",
            accent_bar_color_hex=accents[i % len(accents)],
        )
        time.sleep(0.3)

    controller.set_slide_transition(slide_index=1, transition="fade", speed="fast")
    time.sleep(0.8)

    # ----------------------------------------------------
    # SLIDES 2-5: Individual Detailed Slides for each Product
    # ----------------------------------------------------
    for idx, p in enumerate(products):
        s_num = idx + 2
        print(f">>> Building Slide {s_num}: Detail for {p['name']}...")
        controller.add_slide(layout="blank")
        controller.set_slide_background(s_num, "#F8FAFC")
        time.sleep(0.4)

        # Category Badge
        controller.add_shape(
            slide_index=s_num,
            shape_type="rounded_rectangle",
            left=80,
            top=40,
            width=160,
            height=28,
            fill_color_hex=accents[idx % len(accents)],
            text=p.get("category_name", "Électronique").upper(),
            text_color_hex="#FFFFFF",
            font_size=10,
            bold=True,
        )

        # Title
        controller.add_textbox(
            slide_index=s_num,
            text=p["name"],
            left=80,
            top=75,
            width=800,
            height=50,
            font_size=24,
            bold=True,
            color_hex="#0F172A",
        )

        # Left Column: Price & Overview Card
        overview_text = (
            f"💰 Prix: ${p['price']:.2f} (Économisez ${(p['compare_at_price'] - p['price']):.2f})\n\n"
            f"📦 Disponibilité: {p['stock_quantity']} unités en stock\n\n"
            f"⭐ Évaluation: {p['rating']} / 5 ({p['reviews_count']} avis vérifiés)\n\n"
            f"🏷️ Référence SKU: {p['sku']}\n\n"
            f"📝 Description:\n{p['description']}"
        )
        controller.add_card(
            slide_index=s_num,
            title="Détails Commerciaux",
            body=overview_text,
            left=80,
            top=140,
            width=430,
            height=340,
            bg_color_hex="#FFFFFF",
            border_color_hex="#E2E8F0",
            title_color_hex="#0F172A",
            body_color_hex="#334155",
            accent_bar_color_hex=accents[idx % len(accents)],
        )

        # Right Column: Technical Specifications Card
        specs_lines = [f"• {k}: {v}" for k, v in p.get("attributes", {}).items()]
        specs_body = "\n\n".join(specs_lines) if specs_lines else "Aucune spécification renseignée."
        controller.add_card(
            slide_index=s_num,
            title="Spécifications Techniques",
            body=specs_body,
            left=530,
            top=140,
            width=350,
            height=340,
            bg_color_hex="#FFFFFF",
            border_color_hex="#E2E8F0",
            title_color_hex="#0F172A",
            body_color_hex="#475569",
            accent_bar_color_hex="#64748B",
        )

        # Add speaker notes
        controller.set_speaker_notes(
            slide_index=s_num,
            notes_text=f"Fiche produit pour {p['name']}. SKU: {p['sku']}, Prix: ${p['price']}."
        )

        controller.set_slide_transition(slide_index=s_num, transition="push", speed="fast")
        time.sleep(0.4)

    controller.goto_slide(1)
    print(">>> Finished! 5 beautiful slides generated in PowerPoint live on screen.")


if __name__ == "__main__":
    generate_presentation()
