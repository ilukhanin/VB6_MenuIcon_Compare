# VB6: два способа добавить иконки в стандартное меню

Минимальный проект Visual Basic 6.0. Без OCX, subclassing и owner-draw.

## Где реализация

Всё в `frmMain.frm`, обработчик `Form_Load`:

- **Способ 1**: `SetMenuItemBitmaps(hClassic, i, MF_BYPOSITION, ...)` — помещает монохромную картинку в область галочки меню. Это самый простой вызов. Важно: последние параметры — изображение для *unchecked* и *checked*, а не для hover/normal. В примере они одинаковые.
- **Способ 2**: `SetMenuItemInfo(hModern, i, 1&, itemInfo)` — назначает `MENUITEMINFO.hbmpItem`, с `fMask = MIIM_BITMAP`, не меняя текст пункта. Для наглядности использованы цветные 16×16 BMP (с белым фоном).

`GetMenu` получает верхнее меню формы; `GetSubMenu(..., 0/1)` — конкретный раскрывающийся список; `i` — позиция его пункта от нуля. Иконки загружаются через `LoadPicture` и остаются в `StdPicture` полях формы, чтобы их дескрипторы оставались живы.

## Замечания

- Этот пример добавляет картинки к **элементам выпадающего меню**. Он не вставляет иконки в названия двух пунктов на самой верхней полосе меню.
- `SetMenuItemBitmaps` рассчитан на **монохромные** bitmap; цветные могут выглядеть плохо. Белое в этих bitmap трактуется как прозрачное.
- В `MIIM_BITMAP` здесь обычные 24-битные BMP с **белым фоном**, а не полноценные PNG с альфа-прозрачностью.
- Для обоих способов нужна обычная 32-битная VB6-программа; на 64-битной Windows VB6 EXE запускается через WOW64.
- Проект создан и проверен статически, но **не запускался в Visual Basic 6 IDE** (Windows/VB6 здесь недоступен).

Документация Microsoft:
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setmenuitembitmaps
- https://learn.microsoft.com/en-us/windows/win32/api/winuser/ns-winuser-menuiteminfoa
- https://learn.microsoft.com/en-us/windows/win32/menurc/using-menus
