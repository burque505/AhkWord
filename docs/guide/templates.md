# Templates

## Using a template

**File > New from Template** makes a new document from a template: **Letter**, **Memo**, **Report**, **Court Paper** (with a certificate of service), **Meeting Notes**, or one of your own. A dialog asks for its fields, such as the parties' names; your own details (`{{Your name}}` and the like) are remembered for next time, and date fields get a calendar.

- **Insert > Insert from Template** puts a template's content into the document you are writing. The parts **Caption**, **Signature Block** and **Certificate of Service** are made for that.
- **Insert > Fill In Fields** fills in fields left in a document.

## Making a template

A template is an ordinary document with blanks.

1. **Start from a real document** with the layout you want (a paper you have filed, a letter you have sent), or from a blank one. Whatever AhkWord keeps (fonts, spacing, tabs, indents, tables, header and footer, line numbers), the template keeps.
2. **Replace each part that changes with a field**, typed into the text:

    | Type | The dialog shows |
    |---|---|
    | `{{Plaintiff}}` | a box labelled "Plaintiff" (the name is the label) |
    | `{{Your role\|Plaintiff}}` | the box, with "Plaintiff" filled in if it is left empty |
    | `{{Your address}}` | a box of several lines (any name with "address" in it); each line becomes a paragraph |
    | `{{Hearing date}}`, `{{Date}}` | the box with a calendar beside it (any name with "date" in it): pick a day, or type the date as it should read ("on or about June 3") |

    - The same name used several times is asked for once.
    - Fields whose names begin with "Your" are remembered: your name, address, phone and e-mail fill themselves in next time (the dialog's *Forget remembered values* clears them). Others (the parties, the case) are asked for fresh each time.
    - The filled-in text takes the formatting of the field's first character: format the whole `{{...}}` as the text should look.
    - Use square brackets for text to write by hand, such as `[The first point.]`.

3. **Save it as Rich Text Format into your Templates folder** (`%APPDATA%\AhkWord\Templates`; the start page's *Your templates folder* opens it), named as the template should be called. A piece to insert into other documents, such as a caption or a signature block, goes into the `Parts` folder inside it instead.
4. **Use it** with File > New from Template, or Insert > Insert from Template for a part. To change it later, open the file itself with File > Open, edit it and save it.

Fields in the header and footer are not filled in: they have fields of their own, such as `{page}` (see [Pages and printing](pages.md#headers-and-footers)).
