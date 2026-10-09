"""Production Example: Automated 6-Slide Pitch Deck Generator
Uses the PowerPoint Live Controller MCP suite to generate a complete investor pitch deck.
"""

import os
import sys
import time

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from ppt_controller import controller


def build_pitch_deck():
    print(">>> 1. Initializing Pitch Deck Presentation...")
    controller.new_presentation(width=960, height=540)
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 1: Cover
    # ----------------------------------------------------
    print(">>> Building Slide 1: Cover...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(1, "#0A0F1D", gradient_color_hex="#1E293B")

    controller.add_badge(
        slide_index=1,
        text="SERIES A PITCH DECK",
        left=80,
        top=70,
        bg_color_hex="#1E3A8A",
        text_color_hex="#93C5FD",
        border_color_hex="#3B82F6",
    )

    controller.add_textbox(
        slide_index=1,
        text="Nexus AI : L'Automatisation Bureautique Intelligente",
        left=80,
        top=115,
        width=800,
        height=85,
        font_size=36,
        bold=True,
        color_hex="#F8FAFC",
    )

    controller.add_textbox(
        slide_index=1,
        text="La première plateforme qui connecte les LLMs directement aux logiciels d'entreprise en temps réel.",
        left=80,
        top=210,
        width=780,
        height=50,
        font_size=18,
        color_hex="#94A3B8",
    )
    controller.set_slide_transition(slide_index=1, transition="fade")

    # ----------------------------------------------------
    # SLIDE 2: Problem & Solution
    # ----------------------------------------------------
    print(">>> Building Slide 2: Problem & Solution...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(2, "#F8FAFC")

    controller.add_textbox(
        slide_index=2,
        text="Le Défi & Notre Solution",
        left=80,
        top=45,
        width=800,
        height=45,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )

    controller.add_card(
        slide_index=2,
        title="⚠️ Le Problème",
        body="• Les employés passent plus de 4 heures par semaine à concevoir des diapositives.\n• Les outils actuels d'IA génèrent des fichiers statiques sans retour visuel.\n• Risque constant de décalage de mise en page.",
        left=80,
        top=120,
        width=385,
        height=320,
        bg_color_hex="#FFFFFF",
        border_color_hex="#FCA5A5",
        accent_bar_color_hex="#EF4444",
    )

    controller.add_card(
        slide_index=2,
        title="💡 Notre Solution",
        body="• Intégration Windows COM en direct avec Microsoft PowerPoint.\n• Rendu instantané et visible pendant que l'utilisateur supervise.\n• 31+ outils natifs adaptés aux workflows d'entreprise.",
        left=495,
        top=120,
        width=385,
        height=320,
        bg_color_hex="#FFFFFF",
        border_color_hex="#86EFAC",
        accent_bar_color_hex="#10B981",
    )
    controller.set_slide_transition(slide_index=2, transition="push")

    # ----------------------------------------------------
    # SLIDE 3: Traction & Key Metrics
    # ----------------------------------------------------
    print(">>> Building Slide 3: Metrics...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(3, "#F8FAFC")

    controller.add_textbox(
        slide_index=3,
        text="Traction & Indicateurs Clés de Croissance",
        left=80,
        top=45,
        width=800,
        height=45,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )

    controller.add_metric_card(
        slide_index=3,
        value="$3.2M",
        label="ARR (Revenu Annuel)",
        subtitle="↑ 210% de croissance YoY",
        left=80,
        top=130,
        width=250,
        height=150,
        value_color_hex="#2563EB",
        accent_bar_color_hex="#2563EB",
    )

    controller.add_metric_card(
        slide_index=3,
        value="45,000+",
        label="Utilisateurs Actifs",
        subtitle="Plus de 250 entreprises",
        left=355,
        top=130,
        width=250,
        height=150,
        value_color_hex="#059669",
        accent_bar_color_hex="#059669",
    )

    controller.add_metric_card(
        slide_index=3,
        value="94%",
        label="Rétention NRR",
        subtitle="Adoption virale en équipe",
        left=630,
        top=130,
        width=250,
        height=150,
        value_color_hex="#7C3AED",
        accent_bar_color_hex="#7C3AED",
    )

    # ----------------------------------------------------
    # SLIDE 4: Roadmap
    # ----------------------------------------------------
    print(">>> Building Slide 4: Roadmap...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(4, "#F8FAFC")

    controller.add_textbox(
        slide_index=4,
        text="Feuille de Route & Objectifs 2026-2027",
        left=80,
        top=45,
        width=800,
        height=45,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )

    steps = [
        {"title": "Q1 2026", "desc": "Lancement du serveur MCP PowerPoint avec 30+ outils."},
        {"title": "Q2 2026", "desc": "Intégration d'Excel et Word pour une suite bureautique unifiée."},
        {"title": "Q3 2026", "desc": "Support multi-utilisateurs et collaboration cloud."},
        {"title": "Q4 2026", "desc": "Expansion globale et marketplace de templates."},
    ]

    controller.add_timeline(
        slide_index=4,
        steps=steps,
        left=80,
        top=135,
        width=800,
        height=260,
        accent_color_hex="#2563EB",
    )

    # ----------------------------------------------------
    # SLIDE 5: Testimonial & Quote
    # ----------------------------------------------------
    print(">>> Building Slide 5: Quote...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(5, "#F8FAFC")

    controller.add_textbox(
        slide_index=5,
        text="Ce que disent nos clients",
        left=80,
        top=50,
        width=800,
        height=45,
        font_size=28,
        bold=True,
        color_hex="#0F172A",
    )

    controller.add_quote_card(
        slide_index=5,
        quote="Nexus AI a réduit de 80% le temps de préparation de nos réunions de direction. La visualisation en temps réel change complètement la donne.",
        author="Khadija Benani",
        role="VP Product & Strategy, Global FinTech",
        left=80,
        top=130,
        width=800,
        height=230,
        accent_color_hex="#2563EB",
    )

    # Add Footers to all slides
    controller.add_footer(text="Nexus AI Technologies — Document Confidentiel Série A", show_slide_number=True)

    controller.goto_slide(1)
    print(">>> Finished! 5-Slide Pitch Deck successfully generated live.")


if __name__ == "__main__":
    build_pitch_deck()
