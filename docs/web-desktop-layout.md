# Poliport web desktop layout

User explicitly requested the same structure and visual design as poliport_arsiv.py::_setup_ui. Replace the web dashboard cards with the source application's blue 62px header, grey background, Dosya Ayarları two-row fieldset, expanding Araç Listesi fieldset, İşlem Durumu fieldset with dark log, and blue footer. Use the same table headings, separate driver first/last names, green saved and red pending rows.

Keep browser report URL, host allowlist, history and search accessible in compact optional controls. Keep all download/archive identity, file naming, locking, folder hierarchy and direct .xlsx behavior unchanged. Report file browsing still loads automatically, matching browse_report(), and the explicit Yükle button reloads the chosen file. Browser folder picker can display the selected folder name, not the full local path.

Implementation: update index.html/styles.css and minimal DOM bindings in app.mjs; run existing report, duplicate and direct Excel tests; verify sample report rendering in browser; publish and verify live GitHub Pages.

Validation: 13 existing report/duplicate/direct-Excel/folder tests passed against this checkout. Browser test at native 1250x800 size loaded 4 synthetic rows, confirmed 10 columns including separate driver names and red pending rows, and cleared all selections. No company network requests or real download validation was performed; this change preserves existing download code. Code review identified a stale reload File after drag/drop; fixed by tracking the latest successfully loaded file and routing URL reload to the URL source.
