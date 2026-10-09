"""Test script for advanced PowerPoint utilities:
- Testimonial Quote Cards
- Pros vs Cons Side-by-Side Analysis
- Multi-tier SaaS Pricing Comparison
- Circular Donut Progress Chart
- Global Search and Replace Text
- Professional Document Footers
"""

import time
from ppt_controller import controller


def test_advanced():
    print(">>> 1. Creating presentation for advanced utilities...")
    controller.new_presentation(width=960, height=540)
    time.sleep(1)

    # ----------------------------------------------------
    # SLIDE 1: Testimonial Quote Card
    # ----------------------------------------------------
    print(">>> 2. Building Slide 1 (Quote Card)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(1, "#F8FAFC")

    controller.add_textbox(
        slide_index=1,
        text="Témoignage Client & Avis d'Expert",
        left=80,
        top=50,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    controller.add_quote_card(
        slide_index=1,
        quote="L'intégration de l'automatisation en direct dans PowerPoint a transformé notre productivité. Nos équipes génèrent des présentations complètes en quelques secondes.",
        author="Youssef El Mansouri",
        role="Directeur de l'Innovation, Tech Ventures",
        left=80,
        top=130,
        width=800,
        height=240,
        accent_color_hex="#2563EB",
    )

    # ----------------------------------------------------
    # SLIDE 2: Pros vs Cons (Avantages vs Inconvénients)
    # ----------------------------------------------------
    print(">>> 3. Building Slide 2 (Pros & Cons)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(2, "#F8FAFC")

    controller.add_textbox(
        slide_index=2,
        text="Analyse Stratégique : Pour & Contre",
        left=80,
        top=50,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    pros_list = [
        "Gain de temps massif : création automatisée en quelques secondes",
        "Visualisation directe sur écran pendant que l'IA travaille",
        "Zéro dépendance de bibliothèque externe lourde (COM natif)",
        "Contrôle à distance complet du diaporama et animations",
    ]

    cons_list = [
        "Nécessite Microsoft PowerPoint installé sur la machine hôte",
        "Limité à l'environnement Windows pour les appels COM",
        "Nécessite des permissions d'exécution pour le script local",
    ]

    controller.add_pros_cons(
        slide_index=2,
        pros=pros_list,
        cons=cons_list,
        left=80,
        top=120,
        width=800,
        height=340,
    )

    # ----------------------------------------------------
    # SLIDE 3: Multi-tier Pricing Table
    # ----------------------------------------------------
    print(">>> 4. Building Slide 3 (SaaS Pricing Table)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(3, "#F8FAFC")

    controller.add_textbox(
        slide_index=3,
        text="Grille Tarifaire & Abonnements 2026",
        left=80,
        top=40,
        width=800,
        height=40,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    pricing_tiers = [
        {
            "name": "Starter",
            "price": "$0",
            "period": "Gratuit à vie",
            "highlighted": False,
            "features": [
                "Jusqu'à 10 slides par jour",
                "Export PNG 720p",
                "Support communautaire",
            ]
        },
        {
            "name": "Professionnel",
            "price": "$29",
            "period": "/ utilisateur / mois",
            "highlighted": True,  # Popular!
            "features": [
                "Slides illimitées en direct",
                "Export Full HD 1080p & PDF",
                "Templates KPI, Charts & Code",
                "Support prioritaire 24/7",
            ]
        },
        {
            "name": "Entreprise",
            "price": "$99",
            "period": "/ organisation / mois",
            "highlighted": False,
            "features": [
                "Intégration API & SSO",
                "SLA garanti 99.99%",
                "Templates sur-mesure",
                "Account Manager dédié",
            ]
        }
    ]

    controller.add_pricing_table(
        slide_index=3,
        tiers=pricing_tiers,
        left=80,
        top=110,
        width=800,
        height=370,
    )

    # ----------------------------------------------------
    # SLIDE 4: Donut Progress Chart
    # ----------------------------------------------------
    print(">>> 5. Building Slide 4 (Donut Progress Chart)...")
    controller.add_slide(layout="blank")
    controller.set_slide_background(4, "#F8FAFC")

    controller.add_textbox(
        slide_index=4,
        text="Progression Globale du Projet {{CLIENT_NAME}}",
        left=80,
        top=50,
        width=800,
        height=45,
        font_size=26,
        bold=True,
        color_hex="#0F172A",
    )

    controller.add_donut_chart(
        slide_index=4,
        percentage=84.0,
        label="Taux de Complétion des Tâches (Phase Finale)",
        left=390,
        top=150,
        size=180,
        fill_color_hex="#2563EB",
        bg_color_hex="#F8FAFC",
    )

    # ----------------------------------------------------
    # 6. Test Search & Replace Template Variable
    # ----------------------------------------------------
    print(">>> 6. Testing Search & Replace ({{CLIENT_NAME}} -> 'Acme Corp')...")
    rep_res = controller.search_and_replace_text("{{CLIENT_NAME}}", "Acme Corporation", slide_index=4)
    print("Replace result:", rep_res)

    # ----------------------------------------------------
    # 7. Add Professional Footers to all slides
    # ----------------------------------------------------
    print(">>> 7. Adding Footers to all slides...")
    foot_res = controller.add_footer(text="PowerPoint Live MCP — Document Confidentiel")
    print("Footer result:", foot_res)

    controller.goto_slide(1)
    print(">>> All advanced utilities tested successfully live!")


if __name__ == "__main__":
    test_advanced()
