import { parseXml, type XmlElement } from '../../internal/xml/index.ts';

const decoder = new TextDecoder();
const roots = new WeakMap<Uint8Array, XmlElement>();

/**
 * Shared read-only XML snapshot. Writers must parse their own document; public
 * mutations replace part bytes, naturally invalidating this weakly held root.
 */
export const readPartRoot = (bytes: Uint8Array): XmlElement => {
  let root = roots.get(bytes);
  if (root === undefined) {
    root = parseXml(decoder.decode(bytes)).root;
    roots.set(bytes, root);
  }
  return root;
};
