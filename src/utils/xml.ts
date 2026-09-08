import { XMLParser, XMLValidator } from 'fast-xml-parser';
import { WorkbookError } from './errors';

export type XmlNode = Record<string, unknown>;
export function nodes(value: unknown): XmlNode[] {
  return (Array.isArray(value) ? value : [value]).filter(
    (item): item is XmlNode => typeof item === 'object' && item !== null,
  );
}
export function attribute(node: XmlNode, name: string): string | undefined {
  const value = node[`@_${name}`];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}
export function parseXml(bytes: Uint8Array, rootName: string): XmlNode {
  // OOXML XML parts can be UTF-8 or UTF-16 with a byte order mark.
  const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le'
    : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8';
  let text: string;
  try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); }
  catch { throw new WorkbookError('INVALID_XML', 'An XML part has invalid text encoding.'); }
  if (/<!DOCTYPE|<!ENTITY/i.test(text) || XMLValidator.validate(text) !== true) {
    throw new WorkbookError('INVALID_XML', 'An XML part is malformed or contains a prohibited DTD.');
  }
  const parsed: unknown = new XMLParser({
    ignoreAttributes: false, removeNSPrefix: true,
    parseAttributeValue: false, parseTagValue: false, trimValues: false,
  }).parse(text);
  const root = nodes(parsed)[0]?.[rootName];
  // A self-closing root without attributes is represented as an empty string.
  if (root === '') return {};
  const roots = nodes(root);
  if (roots.length !== 1) throw new WorkbookError('INVALID_XML', `Expected XML root ${rootName}.`);
  return roots[0]!;
}
