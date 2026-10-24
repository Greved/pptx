import { readFileSync } from 'node:fs';
import {
  ParseOption,
  XmlBufferInputProvider,
  XmlDocument,
  XsdValidator,
  xmlCleanupInputProvider,
  xmlRegisterInputProvider,
} from 'libxml2-wasm';

const ecmaRoot = new URL('../../references/ecma-376-5th/ECMA-376/', import.meta.url);

export const schemaFiles = {
  pml: new URL('OfficeOpenXML-XMLSchema-Transitional/pml.xsd', ecmaRoot),
  dml: new URL('OfficeOpenXML-XMLSchema-Transitional/dml-main.xsd', ecmaRoot),
  chart: new URL('OfficeOpenXML-XMLSchema-Transitional/dml-chart.xsd', ecmaRoot),
  rels: new URL('OpenPackagingConventions-XMLSchema/opc-relationships.xsd', ecmaRoot),
  contentTypes: new URL('OpenPackagingConventions-XMLSchema/opc-contentTypes.xsd', ecmaRoot),
  pptxTransitions: new URL('./ms-pptx/transitions.xsd', import.meta.url),
} as const;

export type SchemaKind = keyof typeof schemaFiles;

const parseOptions = {
  option: ParseOption.XML_PARSE_NONET | ParseOption.XML_PARSE_NO_XXE,
};

/** Read the complete graph before compiling: libxml2 can otherwise warn and ignore a missing import. */
export const readSchemaGraph = (entrypoints: readonly URL[]): Record<string, Uint8Array> => {
  const documents: Record<string, Uint8Array> = {};
  const pending = [...entrypoints];
  while (pending.length > 0) {
    const url = pending.pop();
    if (!url || Object.hasOwn(documents, url.href)) continue;
    if (url.protocol !== 'file:') {
      throw new Error(`XML schema infrastructure unavailable: non-local schema ${url.href}`);
    }
    let bytes: Uint8Array;
    try {
      bytes = readFileSync(url);
    } catch (cause) {
      throw new Error(
        `XML schema infrastructure unavailable: ${url.href}; initialize the pinned references/ecma-376-5th submodule`,
        { cause },
      );
    }
    documents[url.href] = bytes;
    const document = XmlDocument.fromBuffer(bytes, { ...parseOptions, url: url.href });
    try {
      const locations = document.find(
        '/xs:schema/xs:import/@schemaLocation | /xs:schema/xs:include/@schemaLocation | /xs:schema/xs:redefine/@schemaLocation',
        { xs: 'http://www.w3.org/2001/XMLSchema' },
      );
      for (const location of locations) pending.push(new URL(location.content, url));
    } finally {
      document.dispose();
    }
  }
  return documents;
};

/** Compile the official schemas with their full import graph, entirely in-process and offline. */
export const compileSchema = (
  entrypoint: URL,
  documents: Record<string, Uint8Array>,
): XsdValidator => {
  const bytes = documents[entrypoint.href];
  if (!bytes) throw new Error(`XML schema infrastructure unavailable: ${entrypoint.href}`);
  if (!xmlRegisterInputProvider(new XmlBufferInputProvider(documents))) {
    throw new Error(
      'XML schema infrastructure unavailable: cannot register the libxml2 schema reader',
    );
  }
  let document: XmlDocument | undefined;
  try {
    document = XmlDocument.fromBuffer(bytes, { ...parseOptions, url: entrypoint.href });
    return XsdValidator.fromDoc(document);
  } catch (cause) {
    throw new Error(`XML schema infrastructure unavailable: cannot compile ${entrypoint.href}`, {
      cause,
    });
  } finally {
    document?.dispose();
    xmlCleanupInputProvider();
  }
};

export const validateXml = (validator: XsdValidator, xml: string): void => {
  const document = XmlDocument.fromString(xml, parseOptions);
  try {
    if (document.dtd) throw new Error('DTD declarations are forbidden in OOXML');
    validator.validate(document);
  } finally {
    document.dispose();
  }
};
