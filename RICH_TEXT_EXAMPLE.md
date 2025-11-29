```tsx
// === Consumer's type definitions (their domain) ===
interface MyAppMetadata {
  richText?: {
    lexical: LexicalEditorState;
  };
  link?: {
    url: string;
  };
}

// === Setup: Type-safe engine ===
const engine = FormulaEngine.buildEmpty<MyAppMetadata>(); // Generic parameter
engine.addWorkbook('wb1');
engine.addSheet({ workbookName: 'wb1', sheetName: 'sheet1' });

const addr = { workbookName: 'wb1', sheetName: 'sheet1', colIndex: 0, rowIndex: 0 };

// === Rich text editing ===
function onUserEditsCell(address, lexicalState: LexicalEditorState) {
  // Store plain text for formulas + rich content for display
  engine.setCellContent(address, extractPlainText(lexicalState));
  engine.setCellMetadata(address, { richText: { lexical: lexicalState } });
  //                                  ^^^^^^^^^ Type-safe! TypeScript knows structure
}

// === Rendering ===
function renderCell(address) {
  const value = engine.getCellValue(address);
  const meta = engine.getCellMetadata(address); // Type: MyAppMetadata | undefined
  const style = engine.getCellStyle(address);
  
  if (value?.startsWith('=')) {
    return <div style={style}>{value}</div>; // Formula - no rich text
  }
  
  if (meta?.richText) {
    return <div style={style}><Lexical state={meta.richText.lexical} /></div>;
  }
  
  if (meta?.link) {
    return <div style={style}><a href={meta.link.url}>{value}</a></div>;
  }
  
  return <div style={style}>{value}</div>;
}

// === Paste ===
// Rich content is automatically copied!
const copiedCells = [addr];
engine.smartPaste(copiedCells, { /* ... */ }, options);
// ✅ Both value AND metadata.richText copied to target

// === Serialization ===
const state = engine.serializeEngine();
// Includes: { metadata: { "A1": { richText: {...} } } }

engine.resetToSerializedEngine(state);
const restored = engine.getCellMetadata(addr);
//    ^^^^^^^^ Type: MyAppMetadata | undefined
```