import { createHash } from 'node:crypto';
import { availableParallelism, freemem, totalmem } from 'node:os';
import { Resvg } from '@resvg/resvg-js';
import { FONT_FILES } from '../src/measure.ts';
import { MONO, SANS, SERIF } from '../src/text-layout.ts';

// A fresh-process native probe, independent of the slide loader and SVG layout.
// The actual raster test reports its own setup/render times beside these phases.
const svg =
  '<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180"><rect width="320" height="180" fill="#3B82F6"/><text font-family="Carlito" font-size="24" x="20" y="40">raster test</text></svg>';
console.info(
  JSON.stringify({
    diagnostic: 'raster-font-environment',
    platform: process.platform,
    arch: process.arch,
    cpus: availableParallelism(),
    totalMemory: totalmem(),
    freeMemory: freemem(),
  }),
);
for (const loadSystemFonts of [false, true, true]) {
  const memory = process.memoryUsage();
  const begin = performance.now();
  const resvg = new Resvg(svg, {
    fitTo: { mode: 'width', value: 320 },
    font: {
      loadSystemFonts,
      fontFiles: FONT_FILES,
      defaultFontFamily: SANS,
      sansSerifFamily: SANS,
      serifFamily: SERIF,
      monospaceFamily: MONO,
    },
  });
  const constructed = performance.now();
  const rendered = resvg.render();
  const rasterized = performance.now();
  const png = rendered.asPng();
  console.info(
    JSON.stringify({
      diagnostic: 'raster-font-native',
      loadSystemFonts,
      constructorMs: constructed - begin,
      renderMs: rasterized - constructed,
      pngMs: performance.now() - rasterized,
      sha256: createHash('sha256').update(png).digest('hex'),
      rssBefore: memory.rss,
      heapBefore: memory.heapUsed,
      externalBefore: memory.external,
    }),
  );
}
