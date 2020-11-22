# KeyValueStore

This module stores KeyValue pairs in the CustomProperties object of an Excel worksheet
If functions are called without a worksheet, the ActiveWorksheet is defaulted to

NOTE: Duplicating a worksheet does not include CustomProperties, i.e. the copy has an empty KeyValueStore again

### "GUI" Macros

1. Macro "KeyValueStore_Init":              Removes all stored key value pairs
2. Macro "KeyValueStore_PasteToTable":      The current KeyValueStore is pasted to an Excel Table at the selection location. The user can modify the values directly.
3. Macro "KeyValueStore_AppendFromTable":   Appends/Updates the KeyValues in the Excel Table back to the KeyValueStore
