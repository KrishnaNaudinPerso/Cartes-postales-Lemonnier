#!/usr/bin/env python3
"""
Script de test rapide — vérifie le groupement des images et la lecture du GED.
Aucun appel API. Affiche un bilan et génère un document Word de test sur 3 cartes.

Usage:
  python3 tester_groupement.py
"""

import sys
import os
from pathlib import Path

# Mock anthropic pour éviter l'erreur d'import dans les tests
class MockAnthropic:
    pass
sys.modules['anthropic'] = MockAnthropic()

# Importer les fonctions du script principal
sys.path.insert(0, str(Path(__file__).parent))
from generer_synthese_cartes import grouper_images, lire_gedcom, DOSSIER

def main():
    print("=" * 55)
    print("  TEST DE GROUPEMENT — Cartes postales Lemonnier")
    print("=" * 55)

    # Test lecture GED
    fichier_ged = next(DOSSIER.glob("*.ged"), None)
    if fichier_ged:
        personnes = lire_gedcom(fichier_ged)
        print(f"\n✅ GED : {len(personnes)} personnes chargées")
        print("\nQuelques personnes :")
        for pid, p in list(personnes.items())[:8]:
            print(f"  • {p['nom_complet']}  naissance:{p['naissance']}  décès:{p['deces']}")
    else:
        print("⚠️ Fichier GED non trouvé")
        personnes = {}

    # Test groupement des images
    cartes = grouper_images(DOSSIER)

    print(f"\n📋 Bilan par carte :")
    print(f"{'N°':>5}  {'Recto':^25}  {'Verso':^25}")
    print("-" * 60)
    for num, info in sorted(cartes.items())[:20]:
        r = info['recto'].name if info['recto'] else "—"
        v = info['verso'].name if info['verso'] else "—"
        print(f"{num:>5}  {r[:25]:<25}  {v[:25]:<25}")
    if len(cartes) > 20:
        print(f"  ... et {len(cartes) - 20} cartes supplémentaires")

    print("\n✅ Test terminé. Lancez maintenant :")
    print("   python3 generer_synthese_cartes.py --test")
    print("   (nécessite la variable ANTHROPIC_API_KEY)")

if __name__ == "__main__":
    main()
