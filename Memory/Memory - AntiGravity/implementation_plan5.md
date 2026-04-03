# Implementation Plan - Add Search Highlighting to DebugForm

Add a search TextBox (`txtDebug`) at the top of `DebugForm` to highlight matching text in `lvwDebug`.

## Proposed Changes

### [DebugForm.vb](file:///d:/Users/Simon/Dropbox/私人文件/Visual%20Studio/Visual%20Studio%2018%20(2026)/Outlook%20Assistant%20-%20(AntiGravity測試區)/DebugForm.vb)

#### 1. Layout Adjustment
In `DebugForm_Load`, enforce control docking:
```vb
txtDebug.Dock = DockStyle.Top
lvwDebug.Dock = DockStyle.Fill
```
This ensures the TextBox is at the top and the ListView fills the rest.

#### 2. Enable OwnerDraw
In `DebugForm_Load`, enable `OwnerDraw`:
```vb
lvwDebug.OwnerDraw = True
AddHandler lvwDebug.DrawColumnHeader, AddressOf lvwDebug_DrawColumnHeader
AddHandler lvwDebug.DrawSubItem, AddressOf lvwDebug_DrawSubItem
AddHandler lvwDebug.DrawItem, AddressOf lvwDebug_DrawItem
```

#### 3. Implement OwnerDraw Handlers
- **DrawColumnHeader**: Set `e.DrawDefault = True`.
- **DrawItem**: Set `e.DrawDefault = True` (or empty, as subitem takes precedence with `OwnerDraw`).
- **DrawSubItem**:
  - If `txtDebug.Text` is not empty and matches subitem text:
    - Draw background (Selected vs Normal).
    - Split text into match and non-match segments.
    - Draw segments with appropriate backgrounds (Yellow for match).
  - Else:
    - Set `e.DrawDefault = True`.

#### 4. Search Event
Handle `txtDebug.TextChanged`:
```vb
Private Sub txtDebug_TextChanged(sender As Object, e As EventArgs) Handles txtDebug.TextChanged
    lvwDebug.Invalidate()
End Sub
```

## Verification Plan

### Manual Verification
1. Run the application.
2. Open the Debug Form (if not already opened).
3. Generate some debug logs.
4. Type in `txtDebug` (e.g., "Begin", "Update", or a specific method name).
5. Verify that:
   - Matching text is highlighted in yellow.
   - Non-matching text remains normal.
   - Selection still works and looks reasonable.
   - Scroll and resize work correctly.
