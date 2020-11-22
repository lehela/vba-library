# VBA Conditional Formatting 

This module's purpose is to save and restore all of a worksheets' Conditional Formats. 

1. The **save** action serializes all Conditional Formats on a worksheet into JSON objects, and persists them in a custom property on the worksheet.

2. The **restore** action clears existing Conditional Formats, fetches the JSON string from the custom property, and deserializes them back into their original state.

The module is still under development, and certainly still contains bugs in certain scenarios. It can probably also be modularized further.

The module uses two other VBA tools:
- VBA-KeyValueStore
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)

