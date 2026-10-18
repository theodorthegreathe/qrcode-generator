# qrcode-generator
pure-VBA module that generates QR codes inside Excel, so it works offline. It needs no add-ins, no internet, no references and no API declarations, and runs in both 32- and 64-bit Office. Solo purpose is generation of SPD codes for usage with various EU/EEA banking apps.


Kazuhiko Arase's MIT-licensed libraries () were used by Claude code as a reference. All the implementations are based on JIS X 0510:1999.

The word "QR Code" is registered trademark of DENSO WAVE INCORPORATED
<br/>http://www.denso-wave.com/qrcode/faqpatent-e.html

# How to use it

In Excel, press Alt+F11, then File → Import File… and choose QRCode.bas. Save the workbook as .xlsm.
Select the cell with your existing =HYPERLINK("https://quickchart.io/qr?text=" & ENCODEURL(A1); …) formula, or the SPD cell A1 itself.
Press Ctrl+Shift+Q (or run the InsertQRCode macro).
The QR picture (100 × 100 pt, about 35 mm) appears right next to the selected cell and moves with the cells.

Your link formula: if you select the hyperlink cell, the macro takes whatever is inside ENCODEURL(...).
Re-running: running the macro again replaces that cell's picture instead of stacking a new one.
Several cells: select several cells to create several codes at once.
Removing codes: DeleteQRCodes macro removes the codes for the selected cells; select the whole sheet to remove all of them.

## Encoding

Mode: typical SPD strings use only uppercase letters, digits and *:.-, so they get the compact alphanumeric mode. Anything else, like lowercase or Czech diacritics in MSG, is encoded as UTF-8 bytes.
Error correction: level M by default, as the QR Platba spec recommends.
Settings: size, error-correction level, and the gap between cell and picture are constants at the top of the module.

Development: importing the module *automatically* needs Excel's "Trust access to the VBA project object model" setting. It is needed only for the automated test runs, manual usage is unaffected by this registry flag.
