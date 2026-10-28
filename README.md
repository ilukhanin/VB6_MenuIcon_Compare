# VB6: Zwei Methoden zum Hinzufügen von Symbolen zu Standardmenüs

Ein minimales Visual Basic 6.0-Projekt. Ohne OCX-Steuerelemente, Subclassing oder Owner-Draw.

## Umsetzung

Der gesamte Code befindet sich in `frmMain.frm`, innerhalb der Ereignisprozedur `Form_Load`:

- **Methode 1**: `SetMenuItemBitmaps(hClassic, i, MF_BYPOSITION, ...)` – platziert ein monochromes Bild im Bereich des Häkchens (Checkmark-Bereich) eines Menüeintrags. Dies ist der einfachste Aufruf. Wichtig: Die letzten Parameter beziehen sich auf den Zustand *nicht markiert* (unchecked) bzw. *markiert* (checked) – nicht auf den Zustand bei Mauszeigerkontakt (Hover) oder im Normalzustand. Im Beispiel sind die Bilder für beide Zustände identisch.
- **Methode 2**: `SetMenuItemInfo(hModern, i, 1&, itemInfo)` – weist `MENUITEMINFO.hbmpItem` ein Bild zu (mit `fMask = MIIM_BITMAP`), ohne den Text des Menüeintrags zu verändern. Zur Veranschaulichung werden hier farbige 16×16-BMP-Dateien (mit weißem Hintergrund) verwendet.

`GetMenu` ruft das Hauptmenü des Formulars ab; `GetSubMenu(..., 0/1)` greift auf das jeweilige Untermenü (Dropdown-Liste) zu; `i` steht für die nullbasierte Position des Menüeintrags. Die Symbole werden mittels `LoadPicture` geladen und in `StdPicture`-Feldern des Formulars gespeichert, damit ihre Handles (Deskriptoren) gültig bleiben.

## Hinweise

- Dieses Beispiel fügt Bilder zu **Elementen in Untermenüs (Dropdown-Menüs)** hinzu. Es werden keine Symbole in die Menüeinträge der obersten Menüleiste selbst eingefügt.
- `SetMenuItemBitmaps` ist für **monochrome** Bitmaps ausgelegt; farbige Bilder können unter Umständen schlecht dargestellt werden. Weiß wird in diesen Bitmaps als transparent interpretiert.
- Bei der Verwendung von `MIIM_BITMAP` kommen hier herkömmliche 24-Bit-BMP-Dateien mit **weißem Hintergrund** zum Einsatz – keine vollwertigen PNG-Dateien mit Alpha-Transparenz.
- Für beide Methoden ist ein reguläres 32-Bit-VB6-Programm erforderlich. Unter 64-Bit-Windows wird die VB6-EXE über WOW64 ausgeführt.
- Das Projekt wurde erstellt und statisch überprüft, jedoch **nicht in der Visual Basic 6-IDE ausgeführt** (Windows/VB6 stehen hier nicht zur Verfügung).

Microsoft-Dokumentation:
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setmenuitembitmaps
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-menuiteminfoa
- https://learn.microsoft.com/en-us/windows/win32/menurc/using-menus
