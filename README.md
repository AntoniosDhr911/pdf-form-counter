# PDF Form Counter

A read-only Windows PowerShell utility for auditing PDF consent forms in school folders.

## Features

- Scan a single school or a batch of school folders.
- Classify PDFs as paper forms, old e-forms, new e-forms, or CDBS e-forms using page counts and naming rules.
- Count unprocessed forms separately, including loose forms under `Forms from t1`.
- Report naming errors, forms in the wrong folders, and forms left unseparated in the entered folder.
- Recognize duplicate folders throughout the hierarchy.
- Export a consolidated timestamped text report for batch scans.

## Requirements

- Windows PowerShell on Windows
- Windows Runtime PDF APIs (`Windows.Data.Pdf`)
- Read access to the folders being scanned

No Python installation is required.

## Usage

1. Download `SchoolFormAuditor.ps1` when available in this repository.
2. Run it from an authorized Windows PowerShell session.
3. Select single-school or batch mode.
4. Select a folder, or resolve a Desktop shortcut to the schools folder.
5. Review the results. Batch mode saves a timestamped TXT report on the Desktop.

## Counting rules

| Type | Classification |
| --- | --- |
| Paper form | Two-page PDF without CDBS in its filename; or a qualifying Consent form folder counted once |
| CDBS e-form | Two-page PDF with CDBS in its filename |
| Old e-form | Five or six pages and smaller than 100 KB |
| New e-form | Seven or eight pages |

The workflow is tailored to a particular folder structure; unusual files may require manual review.

**Duplicate folders:** e-forms inside duplicate folders are excluded from counts. Paper forms inside duplicate folders appear in the Paper forms subtotal but are excluded from Total forms.

**Loose forms:** loose/unprocessed forms are classified and counted, but are not normally flagged for processed-folder naming or placement mistakes.

## Privacy and safety

The scanner must not rename, move, edit, or delete school documents. A batch scan creates a report outside the scanned folders. Reports can contain personal information such as student names, document paths, and school names. **Never publish real reports or school documents.**

## Status

A practical automation script developed for a specific school consent-form workflow. Validate its counts on sample folders before relying on it in another environment.
