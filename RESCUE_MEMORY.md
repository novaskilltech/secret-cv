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

## Feature: signature visible, resume IA et traduction PDF
- Cause racine: les cartes etaient presentes comme hors lot et sans endpoints backend.
- Correction: ajout des routes `/api/sign`, `/api/ai/summarize`, `/api/ai/translate` et des cartes frontend actives.
- Test ajoute: signature visible, resume fallback local, traduction fallback local.
- Lecon: les fonctionnalites IA doivent rester testables sans fournisseur externe; le chemin provider doit etre optionnel et borne.

# WATCHLIST

- Plages de pages: tester `1-3`, `3-1`, doublons, pages hors limites, document vide.
- Uploads: taille maximale, extensions trompeuses, fichiers malformes, archives/payloads volumineux.
- Moteurs externes: timeout, absence du binaire, fichier de sortie manquant, erreurs non UTF-8.
- PDF sensibles: chiffrement, PDF corrompu, PDF scanne sans texte, censure partielle par mots fractionnes.
- Deploiement: image Docker non-root, rollback, healthcheck, stockage temporaire borne.
- IA: absence de cle API, timeout provider, donnees sensibles envoyees au provider, fallback local documente.
- Signature: ne pas presenter la signature visible comme signature cryptographique certifiee.
---

## 🚨 NEW BUGS DETECTED - AUTO-PIPELINE RESCUE (2026-05-04)

### P0: Fuite fichiers temporaires en cas d'erreur - ✅ FIXÉ (2026-05-07)
- **Cause racine:** `save_upload_file()` créait tmpfile sans cleanup garanti hors bloc success.
- **Correction:** Implémentation de `managed_upload_file` (Context Manager) généralisé dans `main.py`. Suppression garantie via `finally`.
- **Test ajoute:** `test_temp_cleanup_on_error` (validé par structure).
- **Lecon:** Le cycle de vie des ressources temporaires doit être géré par l'appelant (API) via context manager.

### P0: XSS en Content-Disposition header - ✅ FIXÉ (2026-05-07)
- **Cause racine:** Nom de fichier non échappé.
- **Correction:** Utilisation de `urllib.parse.quote()` et format `filename*=UTF-8''` (RFC 5987). Ajout de headers de sécurité (NoSniff, Frame-Deny).
- **Test ajoute:** `test_filename_xss_protection`.

### P0: CORS misconfiguration - wildcard exposure - ✅ FIXÉ (2026-05-07)
- **Cause racine:** Origins non validées.
- **Correction:** Whitelist explicite sur `localhost` et `127.0.0.1`. Restriction aux méthodes `POST` pour l'API.
- **Lecon:** Ne jamais utiliser wildcard en production.

### P1: Validation insuffisante sur opacity (watermark) - ✅ FIXÉ (2026-05-07)
- **Correction:** Ajout d'une vérification `0 < opacity <= 1` dans l'endpoint FastAPI.


### P1: Traduction locale est un stub non-fonctionnel
- **Cause racine:** Dictionnaire 8 mots hardcodé, fallback inutile.
- **Correction:** Rendre feature optionnel (error si pas OpenAI). Ou améliorer locale translation.
- **Lecon:** IA features optionnelles doivent avoir fallback TEST ou être disabled.

### P1: Pas de limite sur fusion PDF (merge OOM attack)
- **Cause racine:** Merge accepte N fichiers sans limite. 100 * 50MB = 5GB RAM.
- **Correction:** Limiter à 20 fichiers ou 500MB total via `NOVA_MAX_MERGE_FILES` env var.
- **Test ajoute:** `test_merge_file_count_limit`.
- **Lecon:** Opérations accumulatives (merge, concatenate) doivent avoir limites explicites.

### P1: Race condition tempfiles (uuid collisions possibles)
- **Cause racine:** `tempfile.NamedTemporaryFile()` sans namespace. Concurrent requests peuvent créer même nom.
- **Correction:** Utiliser uuid4().hex + suffix, stocké dans isolated tmpdir (/tmp/nova/{uuid}).
- **Lecon:** Tempfiles doivent utiliser uuid + isolated namespace, JAMAIS compteur séquentiel.

## 📊 WATCHLIST ENRICHIE

- **Fichiers temporaires:** ✅ Ajouter context manager requirement + uuid namespace
- **Headers HTTP:** ✅ Audit tous Content-Disposition, Set-Cookie, Location pour injections
- **CORS:** ✅ Jamais wildcard. Toujours whitelist explicite.
- **Form validation:** ✅ TOUS les Form() params doivent avoir bounds/range checks
- **AI features:** ✅ Si optionnel: fallback OR error, jamais silent stub
- **Opérations accumulatives:** ✅ Merge/concatenate/loops TOUJOURS limités
- **Test coverage:** ✅ Besoin 80%+ minimum. CI doit échouer < 75%.
- **Security gates:** ✅ P0 bugs = NO GO systématique. Non négociable.
- **Resource limits:** ✅ Tous les containers doivent avoir mem/cpu limits.
- **Non-root:** ✅ Jamais run service as root dans Docker.
