# ✉️ Automated Supplier Communications

**Personalized bulk emails to suppliers, straight from Excel: VBA runs a Word Mail Merge for each supplier and sends the result through Outlook, with the supplier's own attachment.**

![Excel VBA](https://img.shields.io/badge/Excel_VBA-217346?style=for-the-badge)
![Word Mail Merge](https://img.shields.io/badge/Word_Mail_Merge-2B579A?style=for-the-badge)
![Outlook](https://img.shields.io/badge/Outlook-0078D4?style=for-the-badge)

![Process flow](docs/process-flow.png)

## The problem

Purchasing teams regularly write to many suppliers: delivery date confirmations, price updates, document requests. Sending these one by one takes hours, and copy-paste mistakes happen. Word's built-in "email merge" cannot add a different attachment per recipient, so it doesn't solve the real case.

## How it works

1. **Excel** holds one row per supplier: contact, email, number of open orders, attachment file name and a `Send` flag.
2. For every supplier with `Send = Yes`, **VBA** runs the **Word Mail Merge** for that single record, so the letter is personalized with Word's formatting.
3. The merged letter becomes the body of a new **Outlook** email, and the supplier's attachment is added.
4. The macro writes the result back to Excel (`Draft created`, `Sent`, `Error: …`) with a timestamp.

**Safe by default:** `Mode = Display` creates drafts you review before sending. `Mode = Send` sends directly and asks for confirmation first.

## Repository structure

```
vba-supplier-mail-merge/
├── src/SupplierMailMerge.bas          # the VBA module
├── workbook/Supplier_Mailing.xlsx     # sample data + settings (import the module, save as .xlsm)
├── template/Supplier_Letter.docx      # Word letter with merge fields
├── attachments/                       # sample supplier attachments (PDF)
└── docs/process-flow.png
```

## How to run it

Requires Windows with Excel, Word and Outlook (desktop).

1. Clone or download the repository.
2. Open `workbook/Supplier_Mailing.xlsx`, press **Alt + F11**, choose **File → Import File** and select `src/SupplierMailMerge.bas`.
3. Save the workbook as **`Supplier_Mailing.xlsm`** (macro-enabled) in the same folder.
4. Check the **Settings** sheet: template path, attachment folder, subject and mode (`Display` or `Send`).
5. Replace the sample emails with test addresses you own, then run the macro **`SendSupplierMails`** (Alt + F8).
6. Review the drafts in Outlook. `ResetStatus` clears the status columns for the next run.

## Settings

| Name | Example | Meaning |
|---|---|---|
| `TemplatePath` | `..\template\Supplier_Letter.docx` | Word letter with `MERGEFIELD`s, relative to the workbook |
| `AttachmentFolder` | `..\attachments` | folder with the files named in the `Attachment` column |
| `Subject` | `Open purchase orders – {SupplierName}` | `{SupplierName}` is replaced per supplier |
| `Mode` | `Display` | `Display` = drafts, `Send` = send immediately |

## Design notes

- **Late binding** (`CreateObject`): no library references, so the file works across Office versions.
- **One record per merge** (`FirstRecord = LastRecord`): the only way to combine Word formatting with a per-supplier attachment.
- **Validation and logging:** invalid emails, missing attachments and errors are written to the `Status` column instead of stopping the run.
- **Clean-up on error:** Word is always closed without saving, even if a run fails.

## What I learned

- Combining Office apps through VBA solves problems that no single app solves alone.
- A "draft first" default is what makes automation safe to hand over to a team.

---

<sub>All supplier names, contacts and emails are fictional (`.example` domains). No company data is included.</sub>
