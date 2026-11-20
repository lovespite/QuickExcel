import upstream from 'papaparse';
export function parse(input, config = {}) {
  if (typeof input !== 'string') throw new TypeError('Papa Parse accepts text in QuickExcel. Read logical files with host.files, or await Blob.text() in an Action.');
  if (config.worker) throw new TypeError('Papa Parse worker mode is unavailable in the isolated QuickExcel worker.');
  if (config.download) throw new TypeError('Use the permission-gated fetch in an Action, then parse the response text.');
  return upstream.parse(input, config);
}
export const { unparse, RECORD_SEP, UNIT_SEP, BYTE_ORDER_MARK, BAD_DELIMITERS, Parser, ParserHandle } = upstream;
export const WORKERS_SUPPORTED = false;
export default { ...upstream, parse, WORKERS_SUPPORTED };
