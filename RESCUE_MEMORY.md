# RESCUE MEMORY

## Bug: reorganisation de pages par plage incomplete
- Cause racine: `parse_page_sequence()` utilisait `range(start_page - 1, end_page - 1, step)`, ce qui excluait toujours la borne finale.
- Correction: calcul de plage en numeros 1-based inclusifs, puis conversion en index 0-based.
- Test ajoute: `test_reorder_range_includes_endpoints` verifie que `3-1` produit bien les pages 3, 2, 1.
- Lecon: toute syntaxe de plage utilisateur doit avoir des tests sur bornes ascendantes et descendantes.

## Bug: uploads non bornes
- Cause racine: `shutil.copyfileobj()` copiait les fichiers sans limite applicative.
- Correction: copie par chunks avec limite `NOVA_MAX_UPLOAD_BYTES`, nettoyage du temporaire en cas de depassement.
- Test ajoute: `test_upload_size_limit_returns_400`.
- Lecon: tout endpoint acceptant un fichier doit appliquer une limite avant le traitement PDF/image/Office.

## Risque: commandes externes sans timeout
- Cause racine: `subprocess.run()` attendait indefiniment LibreOffice/OCRmyPDF en cas de blocage.
- Correction: timeout configurable via `NOVA_COMMAND_TIMEOUT_SECONDS`.
- Test ajoute: couvert indirectement par la suite de non-regression; ajouter un test unitaire mocke si la couche utilitaire grossit.
- Lecon: toute integration moteur externe doit etre bornee par timeout et retour d'erreur controle.

# WATCHLIST

- Plages de pages: tester `1-3`, `3-1`, doublons, pages hors limites, document vide.
- Uploads: taille maximale, extensions trompeuses, fichiers malformes, archives/payloads volumineux.
- Moteurs externes: timeout, absence du binaire, fichier de sortie manquant, erreurs non UTF-8.
- PDF sensibles: chiffrement, PDF corrompu, PDF scanne sans texte, censure partielle par mots fractionnes.
- Deploiement: image Docker non-root, rollback, healthcheck, stockage temporaire borne.
