import { mock, test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request as httpRequest } from 'node:http';
import fs from 'node:fs';
import { syncBuiltinESMExports } from 'node:module';
import { chmod, mkdir, mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { delimiter, join } from 'node:path';
import { resolveCliCommand } from '../src/cli-command.ts';
import { createChat } from '../src/chat.ts';

const pause = () => new Promise((done) => setTimeout(done, 20));

test('CLI resolution launches only recognized Node scripts and keeps arguments separate', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx node cli '));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const env = { PATH: directory };
  const file = join(directory, 'codex');
  for (const shebang of [`#!${process.execPath}`, '#!/usr/bin/env node']) {
    await writeFile(file, `${shebang}\r\nconsole.log('ready');`, { mode: 0o755 });
    assert.deepEqual(resolveCliCommand('codex', env), {
      file: process.execPath,
      args: [file],
    });
  }
  for (const content of ['#!/bin/sh\necho node', '#!/usr/bin/env node --inspect\n', 'node']) {
    await writeFile(file, content);
    assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
  }
  await writeFile(file, '#!/usr/bin/env node' + ' '.repeat(4096) + '\n');
  assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
  await rm(file);
  assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
  await mkdir(file);
  assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
});

test('CLI resolution preserves PATH ordering and POSIX executable permissions', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx-cli-path-'));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const first = join(directory, 'first');
  const second = join(directory, 'second');
  await Promise.all([mkdir(first), mkdir(second)]);
  await writeFile(join(first, 'codex'), '#!/bin/sh\n', { mode: 0o755 });
  await writeFile(join(second, 'codex'), '#!/usr/bin/env node\n', { mode: 0o755 });
  const env = { PATH: [first, second].join(delimiter) };
  assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
  await rm(join(first, 'codex'));
  if (process.platform === 'win32') {
    await writeFile(join(first, 'codex.exe'), 'native executable');
    assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
    await rm(join(first, 'codex.exe'));
  }
  assert.deepEqual(resolveCliCommand('codex', env), {
    file: process.execPath,
    args: [join(second, 'codex')],
  });
  if (process.platform === 'win32') {
    await writeFile(join(directory, 'codex.exe'), 'native executable');
    assert.deepEqual(resolveCliCommand('codex', env, directory), { file: 'codex', args: [] });
  }
  if (process.platform !== 'win32') {
    await chmod(join(second, 'codex'), 0o644);
    assert.deepEqual(resolveCliCommand('codex', env), { file: 'codex', args: [] });
  }
});
test('CLI chat includes focus/history, renders Markdown, serializes edits and stops', async (t) => {
  const directory = await mkdtemp(join(tmpdir(), 'pptx chat '));
  const previousPath = process.env.PATH;
  process.env.PATH = [directory, previousPath].filter(Boolean).join(delimiter);
  const script = `#!${process.execPath}
import {writeFileSync} from 'node:fs';
let input='';for await(const chunk of process.stdin)input+=chunk;
writeFileSync('prompt.txt',input);writeFileSync('args.json',JSON.stringify(process.argv));
if(input.includes('WAIT_FOREVER'))setInterval(()=>{},1000);
else {
 console.log(JSON.stringify({type:'item.completed',item:{type:'agent_message',text:input.includes('MARKDOWN')?'**Bold** and '+String.fromCharCode(96)+'code'+String.fromCharCode(96)+'\\n\\n- item\\n\\n<script>alert(1)</script>\\n\\n[unsafe](javascript:alert(1))':'Codex edit complete'}}));
 console.log(JSON.stringify(input.includes('FAIL_RESULT')?{type:'turn.failed',error:{message:'Result'}}:{type:'turn.completed'}));
}`;
  for (const provider of ['claude', 'codex'])
    await writeFile(join(directory, provider), script, { mode: 0o755 });
  const historyEvents = [];
  let verificationErrors = [];
  let verificationGate;
  const chat = createChat(
    join(directory, 'deck.tsx'),
    () => {},
    () => false,
    async () => {
      await verificationGate;
      return verificationErrors.shift() ?? null;
    },
    undefined,
    undefined,
    {
      async begin() {
        historyEvents.push('begin');
      },
      async end() {
        historyEvents.push('end');
      },
    },
  );
  const focus = {
    slide: 3,
    count: 5,
    revision: 7,
    text: 'Revenue 日本語',
    entry: 'deck.tsx',
    files: ['slides/revenue.tsx'],
    buildError: null,
  };
  const server = createServer(
    (req, res) =>
      void chat.handle(req, res, (index, revision) => {
        assert.equal(index, 2);
        if (revision !== 7) throw new Error('Preview changed');
        return focus;
      }),
  );
  await new Promise((done) => server.listen(0, '127.0.0.1', done));
  const url = `http://127.0.0.1:${server.address().port}`;
  t.after(async () => {
    await chat.close();
    await new Promise((done) => server.close(done));
    process.env.PATH = previousPath;
    await rm(directory, { recursive: true, force: true });
  });
  const post = (body, path = '/chat', origin = url) =>
    fetch(url + path, {
      method: 'POST',
      headers: { origin, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
  const request = (message, provider = 'codex') => ({ message, provider, slide: 2, revision: 7 });
  const state = async () => (await fetch(url + '/chat')).json();
  const finished = async () => {
    for (let i = 0; i < 150; i++) {
      const value = await state();
      if (!value.running) return value;
      await pause();
    }
    throw new Error('Chat did not finish');
  };
  verificationErrors = ['ReferenceError: Bullets is not defined', null];
  await post(request('Create an agenda'));
  let repaired = await finished();
  assert.equal(repaired.status, 'Done');
  assert.deepEqual(historyEvents, ['begin', 'end']);
  assert.equal(repaired.messages.filter((message) => message.role === 'user').length, 2);
  assert.match(await readFile(join(directory, 'prompt.txt'), 'utf8'), /Bullets is not defined/);
  await post({}, '/chat/reset');
  verificationErrors = Array(4).fill('Repeated build failure');
  await post(request('Create an agenda'));
  repaired = await finished();
  assert.match(repaired.status, /after 3 repair attempts/);
  assert.deepEqual(historyEvents, ['begin', 'end', 'begin', 'end']);
  assert.equal(repaired.messages.filter((message) => message.role === 'user').length, 4);
  await post({}, '/chat/reset');
  let releaseVerification;
  verificationGate = new Promise((resolve) => {
    releaseVerification = resolve;
  });
  verificationErrors = ['Do not repair after stop'];
  await post(request('Stop during verification'));
  for (let i = 0; i < 150 && (await state()).status !== 'Verifying preview…'; i++) await pause();
  assert.equal((await state()).status, 'Verifying preview…');
  await post({}, '/chat/stop');
  releaseVerification();
  const stopped = await finished();
  assert.match(stopped.status, /Stopped/);
  assert.deepEqual(historyEvents.slice(-2), ['begin', 'end']);
  assert.equal(stopped.messages.filter((message) => message.role === 'user').length, 1);
  verificationGate = undefined;
  await post({}, '/chat/reset');
  assert.equal((await post(request('cross origin'), '/chat', 'https://example.com')).status, 403);
  assert.equal((await post({ ...request('stale'), revision: 6 })).status, 400);
  // Split inside a Japanese character to exercise HTTP's incremental UTF-8 decoder.
  const encoded = Buffer.from(JSON.stringify(request('この見出しを変更')));
  const split = encoded.indexOf(Buffer.from('こ')) + 1;
  const splitStatus = await new Promise((resolve, reject) => {
    const req = httpRequest(
      url + '/chat',
      {
        method: 'POST',
        headers: { origin: url, 'Content-Type': 'application/json' },
      },
      (res) => {
        res.resume();
        res.on('end', () => resolve(res.statusCode));
      },
    );
    req.on('error', reject);
    req.write(encoded.subarray(0, split));
    setTimeout(() => req.end(encoded.subarray(split)), 30);
  });
  assert.equal(splitStatus, 202);
  let result = await finished();
  assert.equal(result.status, 'Done');
  assert.equal(result.messages[1].text, 'Codex edit complete');
  let prompt = await readFile(join(directory, 'prompt.txt'), 'utf8');
  assert.match(prompt, /Revenue 日本語/);
  assert.match(prompt, /"slide":3/);
  assert.match(prompt, /slides\/revenue.tsx/);
  assert.match(prompt, /NOT a restriction/);
  assert.equal((await post(request('Now edit the entire deck', 'codex'))).status, 202);
  result = await finished();
  assert.equal(result.messages[3].text, 'Codex edit complete');
  prompt = await readFile(join(directory, 'prompt.txt'), 'utf8');
  assert.match(prompt, /この見出しを変更/);
  const args = JSON.parse(await readFile(join(directory, 'args.json'), 'utf8'));
  assert.ok(args.includes('workspace-write'));
  assert.ok(!args.includes('--dangerously-bypass-approvals-and-sandbox'));
  await post(request('WAIT_FOREVER'));
  assert.equal((await post(request('concurrent'))).status, 409);
  assert.equal((await post({}, '/chat/reset')).status, 409);
  await post({}, '/chat/stop');
  assert.match((await finished()).status, /Stopped/);
  await post({}, '/chat/reset');
  assert.deepEqual((await state()).messages, []);
  await post(request('FAIL_RESULT'));
  assert.equal((await finished()).status, 'Result');
  await post(request('MARKDOWN'));
  const rendered = (await finished()).messages.at(-1).html;
  assert.match(rendered, /<strong>Bold<\/strong>/);
  assert.match(rendered, /<code>code<\/code>/);
  assert.ok(!rendered.includes('<script>'));
  assert.ok(!rendered.includes('href="javascript:'));
  assert.equal((await post(request('legacy', 'claude'))).status, 400);
  await rm(join(directory, 'codex'));
  // Keep the real installed CLI out of PATH for this missing-command case.
  process.env.PATH = directory;
  await post(request('missing executable'));
  assert.match((await finished()).status, /Install it and log in/);
  await writeFile(join(directory, 'codex'), script, { mode: 0o755 });
  process.env.PATH = [directory, previousPath].filter(Boolean).join(delimiter);
  await post({}, '/chat/reset');
  const originalStat = fs.statSync;
  const failLaunch = (at) => {
    let resolutions = 0;
    const replacement = mock.method(fs, 'statSync', (...args) => {
      if (args[0] === join(directory, 'codex') && ++resolutions === at)
        throw Object.assign(new Error('CLI filesystem unavailable'), { code: 'EIO' });
      return originalStat(...args);
    });
    syncBuiltinESMExports();
    return () => {
      replacement.mock.restore();
      syncBuiltinESMExports();
    };
  };
  let restoreStat = failLaunch(1);
  try {
    assert.equal((await post(request('startup failure'))).status, 400);
    assert.deepEqual(historyEvents.slice(-2), ['begin', 'end']);
  } finally {
    restoreStat();
  }
  assert.equal((await post(request('retry startup'))).status, 202);
  assert.equal((await finished()).status, 'Done');
  verificationErrors = ['Repair the build'];
  restoreStat = failLaunch(2);
  try {
    assert.equal((await post(request('repair startup failure'))).status, 202);
    assert.match((await finished()).status, /CLI filesystem unavailable/);
    assert.deepEqual(historyEvents.slice(-2), ['begin', 'end']);
  } finally {
    restoreStat();
  }
  await chat.close();
});
