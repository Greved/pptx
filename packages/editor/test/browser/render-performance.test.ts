import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { chromium, type Browser, type Page } from 'playwright';
import {
  addTitleSlide,
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  emu,
  savePresentation,
} from '@office-kit/pptx';
import type { PresentationData, setPresentationTheme } from '@office-kit/pptx';
import type { EditorHandle, mountEditor } from '../../src/index.ts';

declare global {
  interface Window {
    performanceMountEditor: typeof mountEditor;
    performanceEditor: EditorHandle;
    performanceSetTheme: typeof setPresentationTheme;
    performanceCreateDeck: () => PresentationData;
    performanceUpdateFirstText: (pres: PresentationData) => void;
    performanceGlyphNodes: Element[];
    performanceThumbnailGlyphNodes: Element[];
    performanceFitObserves: number;
    performanceExternalCounters: {
      version: number;
      commits: number;
      restores: number;
      changes: number;
    };
  }
}

let browser: Browser;
let script: string;

before(async () => {
  const bundle = await build({
    stdin: {
      contents: `import { mountEditor } from ${JSON.stringify(fileURLToPath(new URL('../../dist/index.js', import.meta.url)))}; import { setPresentationTheme, createPresentation, addTitleSlide, getSlides, getSlideShapes, setShapeText } from '@office-kit/pptx';
        function createDeck() { const pres = createPresentation(); for(let index = 0; index < 60; index++) addTitleSlide(pres, 'Rich text slide ' + (index + 1)); return pres; }
        function updateFirstText(pres) { setShapeText(getSlideShapes(getSlides(pres)[0])[0], 'Changed one object'); }
        Object.assign(window, { performanceMountEditor: mountEditor, performanceSetTheme: setPresentationTheme, performanceCreateDeck: createDeck, performanceUpdateFirstText: updateFirstText });`,
      resolveDir: fileURLToPath(new URL('../..', import.meta.url)),
    },
    bundle: true,
    write: false,
    platform: 'browser',
    format: 'esm',
    conditions: ['browser'],
  });
  const output = bundle.outputFiles[0];
  assert.ok(output);
  script = output.text;
  browser = await chromium.launch({
    headless: true,
    ...(process.env.PLAYWRIGHT_EXECUTABLE_PATH
      ? { executablePath: process.env.PLAYWRIGHT_EXECUTABLE_PATH }
      : {}),
    ...(process.env.PLAYWRIGHT_CHANNEL ? { channel: process.env.PLAYWRIGHT_CHANNEL } : {}),
  });
});

after(async () => {
  await browser?.close();
});

async function observeFitRegistrations(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const NativeResizeObserver = window.ResizeObserver;
    window.performanceFitObserves = 0;
    window.ResizeObserver = class extends NativeResizeObserver {
      override observe(target: Element, options?: ResizeObserverOptions): void {
        if (target.matches('.canvas-area')) window.performanceFitObserves++;
        super.observe(target, options);
      }
    };
  });
}

test('changing one of 40 text objects preserves the other 39 canvas glyph subtrees', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  for (let index = 0; index < 40; index++) {
    addSlideTextBox(slide, {
      x: emu((index % 8) * 1_400_000),
      y: emu(Math.floor(index / 8) * 1_200_000),
      w: emu(1_300_000),
      h: emu(1_000_000),
      text: `Object ${index}: text with wrapping\nSecond paragraph`,
    });
  }
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  await observeFitRegistrations(page);
  await page.route('http://paint.test/**', (route) =>
    route.fulfill({
      contentType: route.request().url().endsWith('/editor.js') ? 'text/javascript' : 'text/html',
      body: route.request().url().endsWith('/editor.js')
        ? script
        : '<!doctype html><style>body{margin:0}#editor{height:950px}</style><div id="editor"></div><script type="module" src="/editor.js"></script>',
    }),
  );
  await page.goto('http://paint.test/');
  await page.waitForFunction(() => typeof window.performanceMountEditor === 'function');
  await page.evaluate(
    async (source) => {
      const target = document.getElementById('editor');
      if (!target) throw new Error('Missing editor host');
      window.performanceEditor = window.performanceMountEditor(target, {
        source: new Uint8Array(source),
        locale: 'en',
      });
      await window.performanceEditor.ready;
      const root = document.querySelector('office-kit-pptx-editor')?.shadowRoot;
      window.performanceGlyphNodes = [
        ...(root?.querySelectorAll('.paint > svg > [data-pptx-shape-id] foreignObject') ?? []),
      ];
    },
    [...(await savePresentation(pres))],
  );
  assert.equal(await page.evaluate(() => window.performanceGlyphNodes.length), 40);
  await page.waitForFunction(
    () =>
      document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelectorAll('.nav [data-slide-index="0"] foreignObject').length === 40,
  );
  await page.evaluate(() => {
    window.performanceThumbnailGlyphNodes = [
      ...(document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelectorAll('.nav [data-slide-index="0"] foreignObject') ?? []),
    ];
  });
  await page
    .locator('.stage .hit')
    .first()
    .click({ position: { x: 1, y: 1 } });
  assert.equal(await page.locator('.panel[hidden] .pane-fields').count(), 0);
  const fitObserves = await page.evaluate(() => window.performanceFitObserves);
  await page.evaluate(async () => {
    await window.performanceEditor.apply('Change one object', window.performanceUpdateFirstText);
  });
  await page.evaluate(() => new Promise<void>((done) => requestAnimationFrame(() => done())));
  assert.match(await page.locator('.paint').innerHTML(), /Changed one object/);
  assert.equal(await page.evaluate(() => window.performanceFitObserves), fitObserves);
  assert.equal(
    await page.evaluate(
      () => window.performanceThumbnailGlyphNodes.filter((node) => node.isConnected).length,
    ),
    39,
  );
  assert.equal(
    await page.evaluate(
      () => window.performanceGlyphNodes.filter((node) => node.isConnected).length,
    ),
    39,
  );
  await page.evaluate(async () => {
    await window.performanceEditor.apply('Change theme', (pres) =>
      window.performanceSetTheme(pres, { dark1: '#1267AB' }),
    );
  });
  await page.waitForFunction(() =>
    document
      .querySelector('office-kit-pptx-editor')
      ?.shadowRoot?.querySelector('.paint')
      ?.innerHTML.toLowerCase()
      .includes('#1267ab'),
  );
  assert.equal(
    await page.evaluate(
      () => window.performanceGlyphNodes.filter((node) => node.isConnected).length,
    ),
    0,
  );
  const stageBeforeResize = await page.locator('.stage').boundingBox();
  assert.ok(stageBeforeResize);
  await page.setViewportSize({ width: 1000, height: 720 });
  await page.waitForFunction((before) => {
    const stage = document
      .querySelector('office-kit-pptx-editor')
      ?.shadowRoot?.querySelector('.stage');
    return stage && Math.abs(stage.getBoundingClientRect().width - before) > 1;
  }, stageBeforeResize.width);
  await page.locator('.stage .hit').first().click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
  const fillSection = page.locator('.pane-section[data-section="fill"]');
  assert.equal(
    await page.locator('.panel:not([hidden]) .pane-section:not([open]) .pane-fields').count(),
    0,
  );
  await fillSection.locator('summary').focus();
  await page.keyboard.press('Enter');
  await fillSection.locator('.pane-fields').waitFor({ state: 'visible' });
  await page.getByRole('button', { name: 'Close Format Shape', exact: true }).click();
  assert.equal(await page.locator('.panel[hidden] .pane-fields').count(), 0);
  await page.locator('.stage .hit').first().click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Format Shape...', exact: true }).click();
  await fillSection.locator('.pane-fields').waitFor({ state: 'visible' });
  await page.getByRole('tab', { name: 'Fill & Line', exact: true }).focus();
  await page.keyboard.press('End');
  assert.equal(
    await page
      .getByRole('tab', { name: 'Size & Properties', exact: true })
      .getAttribute('aria-selected'),
    'true',
  );
  await page.keyboard.press('Home');
  assert.equal(
    await page.getByRole('tab', { name: 'Fill & Line', exact: true }).getAttribute('aria-selected'),
    'true',
  );
  await page.close();
});

test('large decks bound thumbnail SVGs while keeping keyboard rows and live canvas glyphs', async () => {
  const pres = createPresentation();
  for (let index = 0; index < 60; index++) addTitleSlide(pres, `Rich text slide ${index + 1}`);
  const bytes = [...(await savePresentation(pres))];
  const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });
  await observeFitRegistrations(page);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.route('http://performance.test/**', (route) =>
    route.fulfill({
      contentType: route.request().url().endsWith('/editor.js') ? 'text/javascript' : 'text/html',
      body: route.request().url().endsWith('/editor.js')
        ? script
        : '<!doctype html><style>body{margin:0}#editor{height:950px}</style><div id="editor"></div><script type="module" src="/editor.js"></script>',
    }),
  );
  await page.goto('http://performance.test/');
  await page.waitForFunction(() => typeof window.performanceMountEditor === 'function');
  await page.evaluate(async (source) => {
    const target = document.getElementById('editor');
    if (!target) throw new Error('Missing editor host');
    window.performanceEditor = window.performanceMountEditor(target, {
      source: new Uint8Array(source),
      locale: 'en',
    });
    await window.performanceEditor.ready;
  }, bytes);
  const rows = page.locator('.nav .thumb-row');
  await page.locator('.nav .thumb svg').first().waitFor();
  assert.equal(await rows.count(), 60);
  const thumbnailCount = await page.locator('.nav .thumb svg').count();
  assert.ok(
    thumbnailCount > 0 && thumbnailCount < 16,
    `Expected a viewport of previews, got ${thumbnailCount}`,
  );
  await rows.first().focus();
  await page.keyboard.press('End');
  await page.locator('.nav [data-slide-index="59"] .thumb svg').waitFor();
  assert.equal(await rows.last().getAttribute('aria-current'), 'true');
  await page.keyboard.press('Home');
  await page.locator('.nav [data-slide-index="0"] .thumb svg').waitFor();

  const hit = page.locator('.stage .hit').first();
  const gestureFitObserves = await page.evaluate(() => window.performanceFitObserves);
  const box = await hit.boundingBox();
  assert.ok(box);
  const group = page.locator('.paint [data-pptx-shape-id]').first();
  await group.evaluate((element) =>
    element.setAttribute('data-gesture-identity', 'original-glyphs'),
  );
  await page.evaluate(() => document.getSelection()?.removeAllRanges());
  await page.mouse.move(box.x + 1, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + 61, box.y + box.height / 2 + 20, { steps: 5 });
  await page.waitForFunction(() => {
    const root = document.querySelector('office-kit-pptx-editor')?.shadowRoot;
    return root?.querySelector('.paint [data-gesture-identity]')?.hasAttribute('transform');
  });
  assert.equal(await group.getAttribute('data-gesture-identity'), 'original-glyphs');
  assert.match((await group.getAttribute('transform')) ?? '', /^translate\(/);
  assert.equal(
    await page.evaluate(() => document.getSelection()?.type),
    'None',
    'Object movement must not create a native page caret or selection',
  );
  await page.mouse.up();
  await page.waitForFunction(
    () =>
      !document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelector('.paint [data-gesture-identity]'),
  );
  const committedBox = await hit.boundingBox();
  assert.equal(await page.evaluate(() => window.performanceFitObserves), gestureFitObserves);
  assert.ok(committedBox);
  await page.mouse.move(
    committedBox.x + committedBox.width / 4,
    committedBox.y + committedBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    committedBox.x + committedBox.width / 4 + 90,
    committedBox.y + committedBox.height / 2,
    { steps: 3 },
  );
  await page.waitForFunction(() =>
    document
      .querySelector('office-kit-pptx-editor')
      ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
      ?.hasAttribute('transform'),
  );
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await page.waitForFunction(
    () =>
      !document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
        ?.hasAttribute('transform'),
  );
  const cancelledBox = await hit.boundingBox();
  assert.ok(cancelledBox);
  assert.ok(Math.abs(cancelledBox.x - committedBox.x) < 1);
  await hit.dblclick();
  await page.locator('.canvas-shell .inline-edit').waitFor();
  await page.keyboard.press('Control+A');
  assert.match(
    await page.evaluate(() => document.getSelection()?.toString() ?? ''),
    /Rich text slide/,
  );
  await page.keyboard.type('Updated rich text');
  const draftBox = await page.locator('.stage .hit').nth(1).boundingBox();
  assert.ok(draftBox);
  await page.mouse.move(draftBox.x + draftBox.width / 4, draftBox.y + 1);
  await page.mouse.down();
  await page.mouse.move(draftBox.x + draftBox.width / 4 + 60, draftBox.y + 31, { steps: 3 });
  await page.waitForFunction(
    () =>
      !!document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id][transform]'),
  );
  await page.mouse.up();
  await page.locator('.canvas-shell .inline-edit').waitFor({ state: 'hidden' });
  await page.waitForFunction(() => {
    const root = document.querySelector('office-kit-pptx-editor')?.shadowRoot;
    return root
      ?.querySelector('.nav [data-slide-index="0"] .thumb')
      ?.textContent?.includes('Updated rich text');
  });
  await page.evaluate(async () => {
    await window.performanceEditor.apply('Change theme', (presentation) => {
      window.performanceSetTheme(presentation, { dark1: '#1267AB' });
    });
  });
  await page.waitForFunction(() => {
    const root = document.querySelector('office-kit-pptx-editor')?.shadowRoot;
    return root
      ?.querySelector('.nav [data-slide-index="0"] .thumb')
      ?.innerHTML.toLowerCase()
      .includes('#1267ab');
  });
  await page.evaluate(async (source) => {
    window.performanceEditor.destroy();
    const target = document.getElementById('editor');
    if (!target) throw new Error('Missing editor host');
    const counters = { version: 0, commits: 0, restores: 0, changes: 0 };
    window.performanceExternalCounters = counters;
    window.performanceEditor = window.performanceMountEditor(target, {
      source: new Uint8Array(source),
      locale: 'en',
      externalHistory: {
        reset() {},
        commit() {
          counters.version++;
          counters.commits++;
        },
        restore() {
          counters.restores++;
          return window.performanceCreateDeck();
        },
        canUndo: false,
        canRedo: false,
        undoLabel: null,
        redoLabel: null,
      },
    });
    await window.performanceEditor.ready;
    window.performanceEditor.on('change', () => {
      counters.changes++;
    });
  }, bytes);
  const externalBox = await hit.boundingBox();
  assert.ok(externalBox);
  await page.mouse.move(
    externalBox.x + externalBox.width / 4,
    externalBox.y + externalBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    externalBox.x + externalBox.width / 4 + 60,
    externalBox.y + externalBox.height / 2,
    { steps: 3 },
  );
  await page.waitForFunction(() =>
    document
      .querySelector('office-kit-pptx-editor')
      ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
      ?.hasAttribute('transform'),
  );
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await page.waitForFunction(
    () =>
      !document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
        ?.hasAttribute('transform'),
  );
  assert.deepEqual(await page.evaluate(() => window.performanceExternalCounters), {
    version: 0,
    commits: 0,
    restores: 1,
    changes: 0,
  });
  assert.equal(await page.evaluate(() => window.performanceEditor.dirty), false);
  const restoredExternalBox = await hit.boundingBox();
  assert.ok(restoredExternalBox);
  assert.ok(Math.abs(restoredExternalBox.x - externalBox.x) < 1);
  await hit.click({ position: { x: 1, y: 10 } });
  const rotate = await page.locator('.stage .rotate').boundingBox();
  assert.ok(rotate);
  const beforeRotation = await group.boundingBox();
  assert.ok(beforeRotation);
  await group.evaluate((element) =>
    element.setAttribute('data-gesture-identity', 'rotation-glyphs'),
  );
  await page.mouse.move(rotate.x + rotate.width / 2, rotate.y + rotate.height / 2);
  await page.mouse.down();
  await page.mouse.move(rotate.x + rotate.width / 2 + 90, rotate.y + rotate.height / 2 + 40, {
    steps: 3,
  });
  await page.waitForFunction(() =>
    document
      .querySelector('office-kit-pptx-editor')
      ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
      ?.getAttribute('transform')
      ?.startsWith('rotate('),
  );
  assert.equal(await group.getAttribute('data-gesture-identity'), 'rotation-glyphs');
  const duringRotation = await group.boundingBox();
  assert.ok(duringRotation);
  assert.ok(
    Math.abs(duringRotation.width - beforeRotation.width) > 1 ||
      Math.abs(duringRotation.height - beforeRotation.height) > 1,
  );
  await page.keyboard.press('Escape');
  await page.mouse.up();
  await page.waitForFunction(
    () =>
      !document
        .querySelector('office-kit-pptx-editor')
        ?.shadowRoot?.querySelector('.paint [data-pptx-shape-id]')
        ?.hasAttribute('transform'),
  );
  assert.deepEqual(await page.evaluate(() => window.performanceExternalCounters), {
    version: 0,
    commits: 0,
    restores: 2,
    changes: 0,
  });
  assert.deepEqual(errors, []);
  await page.close();
});
