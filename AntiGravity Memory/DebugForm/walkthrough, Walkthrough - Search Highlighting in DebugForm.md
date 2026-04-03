# Walkthrough - Search Highlighting in DebugForm

I have implemented the search layout and text highlighting functionality in `DebugForm`.

## Changes Made

### Layout & initialization
- In `DebugForm_Load`, added code to dock `txtDebug` to the Top and `lvwDebug` to Fill.
- Enabled `OwnerDraw` for `lvwDebug` so we can customize rendering.

### Search Event
- Added `txtDebug_TextChanged` event handler to invalidate and trigger redraw of `lvwDebug` whenever the search string changes.

### Custom Drawing (OwnerDraw)
- **DrawColumnHeader**: Keeps default behavior (`e.DrawDefault = True`).
- **DrawItem**: Keeps default behavior for Details view pass-through (`e.DrawDefault = True`).
- **DrawSubItem**:
  - Checks if text matches the search phrase.
  - If match found, renders with custom highlight sequence:
    1. Background (Selection Color or default Item BackColor).
    2. Split items to render non-match and match chunks.
    3. Handles text alignment (Left, Center, Right) as configured on each Column node.
    4. Highlight matching chunks with Yellow background and Black text.
  - Reverts to `e.DrawDefault = True` for rows/nodes not matching standard conditions.

## Verification checklist
To verify the implementation:
1. [ ] Run the app forming logs with debugging active.
2. [ ] Type a phrase to see if matched terms render bright yellow inside respective boxes.
