import { afterAll } from 'vitest';
import type { XsdValidator } from 'libxml2-wasm';
import { compileSchema, readSchemaGraph, schemaFiles, validateXml } from './schema-validator.ts';
import type { SchemaKind } from './schema-validator.ts';

export type { SchemaKind } from './schema-validator.ts';

// A missing schema/import fails collection instead of silently skipping validation tests.
const documents = readSchemaGraph(Object.values(schemaFiles));
const validators = new Map<SchemaKind, XsdValidator>();
afterAll(() => {
  for (const validator of validators.values()) validator.dispose();
  validators.clear();
});

const validatorFor = (kind: SchemaKind): XsdValidator => {
  const cached = validators.get(kind);
  if (cached) return cached;
  const validator = compileSchema(schemaFiles[kind], documents);
  validators.set(kind, validator);
  return validator;
};

/** Kept for existing callers; required infrastructure now throws instead of returning false. */
export const isSchemaValidationAvailable = (): true => {
  validateXml(
    validatorFor('rels'),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
  );
  return true;
};

export const expectSchemaValid = (xml: string, kind: SchemaKind): void => {
  const validator = validatorFor(kind);
  try {
    validateXml(validator, xml);
  } catch (cause) {
    const message = cause instanceof Error ? cause.message : String(cause);
    throw new Error(`schema validation failed (${kind}):\n${message}`, { cause });
  }
};

export const expectSchemaValidAll = (
  documents: ReadonlyArray<{ readonly name: string; readonly xml: string }>,
  kind: SchemaKind,
): void => {
  for (const document of documents) {
    try {
      expectSchemaValid(document.xml, kind);
    } catch (cause) {
      const message = cause instanceof Error ? cause.message : String(cause);
      throw new Error(`${document.name}: ${message}`, { cause });
    }
  }
};
