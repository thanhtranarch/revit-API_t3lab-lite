# Revit API Best Practices
- **Transactions:** Always wrap model changes in `with revit.Transaction("Command Name"):`.
- **Performance:** Use `FilteredElementCollector` effectively. Avoid nested loops for element filtering.
- **Safety:** Verify `doc.IsFamilyDocument` or `doc.IsReadOnly` before execution.
- **UI:** Use `pyrevit.forms` for user interaction (alerts, selectors).
