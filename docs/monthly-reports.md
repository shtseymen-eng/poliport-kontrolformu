# Monthly report capacity fix

Observed error comes from core.mjs entry(): any decompressed XML entry above 50 MiB is rejected. Uploads are also capped at 20 MiB and URL downloads at 25 MiB. Full-sheet DOM plus materializing every row multiplies memory usage, so merely lifting the cap is insufficient.

Plan: regression test with a valid 31-day worksheet above 50 MiB; stream ZIP entry chunks through pinned sax 1.6.1, keeping only Terminal records and the necessary hyperlinks; run parsing in a cancellable dedicated Worker; show read progress; raise report file and report URL limits to 512 MiB, worksheet streaming limit to 2 GiB and shared strings to 512 MiB. Individual downloaded control forms keep their existing 25 MiB cap. Keep all dates, deduplication keys, folder paths and file names unchanged. Update Program Merkezi catalog revision after deployment.

Tests: 31-day large XML, existing report equivalence, relationships, shared strings/1904 dates, malformed XML, explicit stream byte cap, Worker cancellation/error/transfer behavior, browser sample import and published asset checks. Actual user workbook was not provided; distinguish synthetic verification from a real company workbook test.

Verified: 10 new regression/Worker/folder-cancellation checks plus 13 existing date/identity/folder/direct-Excel checks passed. Browser Worker imported a synthetic 31-day workbook with 31,000 rows (worksheet above 50 MiB; compressed workbook above old 25 MiB fetch limit), showed responsive progress and 3,100 matching Terminal records across 31 result pages. Folder index checks now skip fresh folders, deduplicate keys, run at bounded concurrency and stage results until completion; stopping preserves the previous list. No actual company workbook was provided.
