import crypto from 'node:crypto';

// PDF.js 6.5.438, Apache-2.0. See patches/pdfjs-actualtext.md.
// Only the text extractor changes; page painting and the exported PDF do not.
export function patchPdfWorker(source) {
  const upstreamHash = '43710009ea26bfbcb615524946e46e3aaf8ad927514f5e465e8ff6d18044ffe9';
  if (crypto.createHash('sha256').update(source).digest('hex') !== upstreamHash)
    throw new Error('PDF.js worker changed; review the ActualText patch before upgrading');
  const start = source.indexOf('  async getTextContent({');
  const end = source.indexOf('  async extractDataStructures(', start);
  let extractor = source.slice(start, end);
  function replace(before, after) {
    if (extractor.split(before).length !== 2) throw new Error('ActualText patch context is not unique');
    extractor = extractor.replace(before, after);
  }
  replace('    let markedContentLevel = 0;', `    // QuickExcel: Skia emits ActualText for glyph aliases/ligatures.
    // Keep frames across streaming chunks. Nested tags inherit the outer replacement.
    const actualTextFrames = [];
    function beginActualText(value) {
      if (actualTextFrames.length >= 64) throw new FormatError("Marked content nesting exceeds preview budget");
      actualTextFrames.push(typeof value === "string" ? { text: stringToPDFString(value), emitted: false } : null);
    }
    let markedContentLevel = 0;`);
  replace('      textContent.items.push(runBidiTransform(textContentItem));', `      const actual = actualTextFrames.find(frame => frame !== null);
      if (actual) {
        // Each Skia span is one shaped cluster. Emit its original code points once,
        // retaining the glyph positioning for selection and avoiding global NFKC.
        textContentItem.str = actual.emitted ? [] : [actual.text];
        actual.emitted = true;
      }
      textContent.items.push(runBidiTransform(textContentItem));`);
  replace('          case OPS.beginMarkedContent:\n            flushTextContentItem();',
    '          case OPS.beginMarkedContent:\n            flushTextContentItem();\n            beginActualText(null);');
  replace('          case OPS.beginMarkedContentProps:\n            flushTextContentItem();',
    '          case OPS.beginMarkedContentProps:\n            flushTextContentItem();\n            beginActualText(args[1] instanceof Dict ? args[1].get("ActualText") : null);');
  replace('          case OPS.endMarkedContent:\n            flushTextContentItem();',
    '          case OPS.endMarkedContent:\n            flushTextContentItem();\n            actualTextFrames.pop();');
  return source.slice(0, start) + extractor + source.slice(end);
}
