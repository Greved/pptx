import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { describe, expect, it } from 'vitest';
import { xmlCleanupInputProvider, xmlRegisterInputProvider } from 'libxml2-wasm';
import {
  expectSchemaValid,
  expectSchemaValidAll,
  isSchemaValidationAvailable,
} from './lib/expect-schema-valid.ts';
import {
  compileSchema,
  readSchemaGraph,
  schemaFiles,
  validateXml,
} from './lib/schema-validator.ts';

const validRelationships =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

describe('required portable XML schema validator', () => {
  it('runs without an installed xmllint executable', () => {
    expect(isSchemaValidationAvailable()).toBe(true);
    expect(() => expectSchemaValid(validRelationships, 'rels')).not.toThrow();
  });

  it('rejects malformed XML', () => {
    expect(() => expectSchemaValid('<Relationships>', 'rels')).toThrow(/schema validation failed/);
  });

  it('rejects well-formed XML that violates the schema', () => {
    expect(() => expectSchemaValid('<not-pml/>', 'pml')).toThrow(/schema validation failed/);
  });

  it('names the invalid document in a batch', () => {
    expect(() =>
      expectSchemaValidAll(
        [
          { name: 'valid.rels', xml: validRelationships },
          { name: 'invalid.rels', xml: '<not-relationships/>' },
        ],
        'rels',
      ),
    ).toThrow(/invalid\.rels/);
  });

  it('enforces imported DrawingML types in PresentationML', () => {
    const slide = (color: string): string =>
      `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`;
    expect(() => expectSchemaValid(slide('FF0000'), 'pml')).not.toThrow();
    expect(() => expectSchemaValid(slide('not-a-color'), 'pml')).toThrow(
      /schema validation failed/,
    );
  });

  it('fails closed when the schema tree is missing', () => {
    const missing = pathToFileURL(join(tmpdir(), 'office-kit-missing-schema', 'root.xsd'));
    expect(() => readSchemaGraph([missing])).toThrow(/XML schema infrastructure unavailable/);
  });

  it('fails closed when a transitive schema import is missing', () => {
    const directory = mkdtempSync(join(tmpdir(), 'office-kit-schema-import-'));
    const root = join(directory, 'root.xsd');
    try {
      writeFileSync(
        root,
        '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:import namespace="urn:missing" schemaLocation="missing.xsd"/></xs:schema>',
      );
      expect(() => readSchemaGraph([pathToFileURL(root)])).toThrow(/missing\.xsd/);
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  });

  it('rejects network schema imports instead of fetching or ignoring them', () => {
    expect(() => readSchemaGraph([new URL('https://example.invalid/schema.xsd')])).toThrow(
      /non-local schema/,
    );
  });

  it('fails closed when the schema compiler cannot initialize', () => {
    const url = new URL('file:///invalid-schema.xsd');
    expect(() =>
      compileSchema(url, {
        [url.href]: new TextEncoder().encode('<not-a-schema/>'),
      }),
    ).toThrow(/XML schema infrastructure unavailable: cannot compile/);
  });

  it.each([
    '<!DOCTYPE Relationships SYSTEM "file:///private-secret.dtd">',
    '<!DOCTYPE Relationships SYSTEM "https://example.invalid/private.dtd">',
    '<!DOCTYPE Relationships [<!ENTITY secret SYSTEM "file:///private-secret.txt">]>',
  ])('rejects DTDs and external entities without accessing their targets: %s', (declaration) => {
    const validator = compileSchema(schemaFiles.rels, readSchemaGraph([schemaFiles.rels]));
    const externalReads: string[] = [];
    expect(
      xmlRegisterInputProvider({
        match: (url) => {
          externalReads.push(url);
          return true;
        },
        open: () => undefined,
        read: () => -1,
        close: () => true,
      }),
    ).toBe(true);
    try {
      expect(() => validateXml(validator, declaration + validRelationships)).toThrow(
        'DTD declarations are forbidden in OOXML',
      );
      expect(externalReads).toEqual([]);
    } finally {
      xmlCleanupInputProvider();
      validator.dispose();
    }
  });
});
