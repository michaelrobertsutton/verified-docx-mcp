# #45 — formatting (merge gate)

Depends on #42. Keep draft until tested on exact stack commits in disposable
local and SharePoint documents. Record Word version, requirement sets and font
read-back evidence. Supported preservation fields: bold, italic, underline,
strike/double-strike, color, font name/size, subscript and superscript.

Replace the first plain sentence immediately after a bold lead-in: replaced
(default) must preserve plain formatting, previous must retain host inheritance,
none must remove direct formatting (WordApiDesktop 1.3 required). Mixed-format
replaced matches must refuse before writing. Test multiple matches, empty
replacement, and tracking-mode restoration on failure. Apply every format_text
property to a table-cell substring and check independent read-back and Word UI.
A protected range that refuses formatting must not return verified success.
