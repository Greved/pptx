import assert from 'node:assert/strict';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { chromium } from 'playwright';
import type { svgPaint } from '../../src/canvas/svg-paint.ts';

declare global {
  interface Window {
    testSvgPaint: typeof svgPaint;
  }
}

test('SVG paint preserves exact structure, namespaces, links and unchanged root nodes', async () => {
  const bundle = await build({
    stdin: {
      contents: `import { svgPaint } from ${JSON.stringify(fileURLToPath(new URL('../../src/canvas/svg-paint.ts', import.meta.url)))}; window.testSvgPaint = svgPaint;`,
      resolveDir: fileURLToPath(new URL('../..', import.meta.url)),
    },
    bundle: true,
    write: false,
    platform: 'browser',
    format: 'iife',
  });
  const output = bundle.outputFiles[0];
  assert.ok(output);
  const browser = await chromium.launch({
    headless: true,
    ...(process.env.PLAYWRIGHT_CHANNEL ? { channel: process.env.PLAYWRIGHT_CHANNEL } : {}),
    ...(process.env.PLAYWRIGHT_EXECUTABLE_PATH
      ? { executablePath: process.env.PLAYWRIGHT_EXECUTABLE_PATH }
      : {}),
  });
  try {
    const page = await browser.newPage();
    await page.setContent('<!doctype html><div id="paint"></div>');
    await page.addScriptTag({ content: output.text });
    const incremental = await page.evaluate(() => {
      const node = document.createElement('div');
      document.body.append(node);
      const opening = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">';
      const fragments = Array.from(
        { length: 40 },
        (_, index) =>
          `<g data-pptx-shape-id="${index}"><foreignObject><div xmlns="http://www.w3.org/1999/xhtml">Text ${index}</div></foreignObject></g>`,
      );
      const result = (parts: string[]) => ({
        dependsOnBrowserFontMetrics: false,
        opening,
        fragments: parts,
        svg: opening + parts.join('') + '</svg>',
      });
      const action = window.testSvgPaint(node, result(fragments));
      const original = [...node.querySelectorAll('foreignObject')];
      const descriptor = Object.getOwnPropertyDescriptor(Element.prototype, 'innerHTML');
      if (!descriptor?.set) throw new Error('Missing innerHTML setter');
      const setter = descriptor.set;
      const parsed: string[] = [];
      Object.defineProperty(Element.prototype, 'innerHTML', {
        ...descriptor,
        set(value: string) {
          parsed.push(String(value));
          setter.call(this, value);
        },
      });
      try {
        const changed = [...fragments];
        changed[0] = '<g data-pptx-shape-id="0"><text>Changed</text></g>';
        action.update(result(changed));
        return {
          parsed: parsed.length,
          parsedUnchanged: parsed.some((value) => value.includes('Text 1')),
          retained: original.filter((element) => element.isConnected).length,
          text: node.querySelector('text')?.textContent,
        };
      } finally {
        Object.defineProperty(Element.prototype, 'innerHTML', descriptor);
        action.destroy();
        node.remove();
      }
    });
    assert.deepEqual(incremental, {
      parsed: 1,
      parsedUnchanged: false,
      retained: 39,
      text: 'Changed',
    });
    const result = await page.evaluate(() => {
      const node = document.getElementById('paint');
      if (!(node instanceof HTMLDivElement)) throw new Error('Missing paint container');
      const root = (children: string, width = 100) =>
        `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 ${width} 100">${children}</svg>`;
      const defs =
        '<defs><linearGradient id="fill"><stop stop-color="red"/></linearGradient></defs>';
      const background = '<rect width="100" height="100" fill="white"/>';
      const first = '<g data-pptx-shape-id="1"><a href="#first"><text>First</text></a></g>';
      const changed = '<g data-pptx-shape-id="1"><a href="#changed"><text>Changed</text></a></g>';
      const second =
        '<g data-pptx-shape-id="2"><g transform="rotate(10)"><use xlink:href="#fill"/><a href="#second" tabindex="0"><text>Second</text></a></g></g>';
      const action = window.testSvgPaint(node, root(defs + background + first + second));
      const originalSvg = node.firstElementChild;
      const originalSecond = node.querySelector('[data-pptx-shape-id="2"]');
      const preservedLink = originalSecond?.querySelector('a');
      if (!(preservedLink instanceof SVGElement)) throw new Error('Missing SVG link');
      preservedLink.focus();
      action.update(root(defs + background + changed + second));
      const focusPreserved = document.activeElement === preservedLink;
      const reused = node.querySelector('[data-pptx-shape-id="2"]') === originalSecond;
      const svgReused = node.firstElementChild === originalSvg;
      const link = node.querySelector('a')?.getAttribute('href');
      const namespace = node
        .querySelector('use')
        ?.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
      const reordered = root(defs + background + second + changed);
      action.update(reordered);
      const expected = document.createElement('template');
      expected.innerHTML = reordered;
      const reorderExact = node.firstElementChild?.isEqualNode(expected.content.firstElementChild);
      action.update(root(defs + background + second));
      const removed = !node.querySelector('[data-pptx-shape-id="1"]');
      action.update(root(defs + background + second, 200));
      const resize =
        node.firstElementChild !== originalSvg &&
        node.firstElementChild?.getAttribute('viewBox') === '0 0 200 100';
      action.update('<!-- render error -->');
      const fallback = node.innerHTML === '<!-- render error -->';
      action.update(root(second));
      const recovered = node.querySelector('svg')?.namespaceURI === 'http://www.w3.org/2000/svg';
      action.destroy();
      return {
        reused,
        svgReused,
        focusPreserved,
        link,
        namespace,
        reorderExact,
        removed,
        resize,
        fallback,
        recovered,
        destroyed: node.childNodes.length === 0,
      };
    });
    assert.deepEqual(result, {
      reused: true,
      svgReused: true,
      focusPreserved: true,
      link: '#changed',
      namespace: '#fill',
      reorderExact: true,
      removed: true,
      resize: true,
      fallback: true,
      recovered: true,
      destroyed: true,
    });
  } finally {
    await browser.close();
  }
});
